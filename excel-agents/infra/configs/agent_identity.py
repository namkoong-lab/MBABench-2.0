"""Resolve a run's full agent settings from its cohort label
(task_attempts.agent_model_name) via the agent_identities.yaml registry.

A run config names its cohort with `agent_model_name`; nothing else about
the model may appear in the config. The registry entry for that label
supplies every setting that changes what the agent does:

    provider, ui_model_label, thinking_effort, agent_folder, agent_model_type

The runner injects them into the engine config, the engine selects AND
verifies the pinned UI state, and the DB row records the same dict in
task_attempts.extra_configs. An unknown label refuses to run and prints a
paste-ready stanza; a config that sets any pinned key refuses to run too,
so two rows under one label can never have run with different settings.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any, NamedTuple

import yaml

REGISTRY_PATH = Path(__file__).resolve().parents[2] / "agent_identities.yaml"

# One combination, one cohort.
AxisKey = tuple[str, str | None, str | None]

VALID_PROVIDERS = ("claude_excel_agent", "chatgpt_excel_agent")

# Every setting an identity pins, in stanza order. (key, type, nullable).
# A run config must not set any of these (they live only in the registry).
PINNED_KEYS = (
    ("provider", str, False),
    ("ui_model_label", str, True),
    ("thinking_effort", str, True),
    ("agent_folder", str, False),
    ("agent_model_type", str, False),
)
PINNED_KEY_NAMES = tuple(k for k, _, _ in PINNED_KEYS)

# Keys a run config might try to sneak into a provider block to steer the
# model — refused because the registry is the only source for them.
_PROVIDER_BLOCK_PINNED = ("model", "ui_model_label", "thinking_effort")


class AgentIdentity(NamedTuple):
    agent_model_name: str
    provider: str
    ui_model_label: str | None
    thinking_effort: str | None
    agent_folder: str
    agent_model_type: str

    # task_io compatibility (mirrors gui's AgentIdentity attribute names).
    @property
    def model_name(self) -> str:
        return self.agent_model_name

    @property
    def axes(self) -> AxisKey:
        return (self.provider, self.ui_model_label, self.thinking_effort)

    def settings(self) -> dict[str, Any]:
        """The pinned keys, as injected into the engine config and stored
        in task_attempts.extra_configs."""
        return {k: getattr(self, k) for k in PINNED_KEY_NAMES}

    extra_configs = settings


class AgentIdentityError(ValueError):
    """Registry is invalid, the label is unknown, or the config sets a pinned key."""


def _entry_to_identity(entry: dict[str, Any], where: str) -> AgentIdentity:
    label = entry.get("agent_model_name")
    if not label:
        raise AgentIdentityError(f"{where}: entry is missing 'agent_model_name'")
    values: dict[str, Any] = {}
    for key, typ, nullable in PINNED_KEYS:
        if key not in entry:
            raise AgentIdentityError(
                f"{where}: entry is missing '{key}' — it changes what the agent "
                f"does, so every identity must pin it"
                + (" (use null if the provider does not take it)" if nullable else "")
            )
        value = entry[key]
        if value is None:
            if not nullable:
                raise AgentIdentityError(f"{where}: '{key}' may not be null")
        elif not isinstance(value, typ):
            raise AgentIdentityError(
                f"{where}: '{key}' must be {typ.__name__}, got {value!r}"
            )
        values[key] = value
    identity = AgentIdentity(str(label), **values)
    if identity.provider not in VALID_PROVIDERS:
        raise AgentIdentityError(
            f"{where}: provider {identity.provider!r} is not one of "
            f"{VALID_PROVIDERS}"
        )
    if identity.agent_model_type != "excel":
        raise AgentIdentityError(
            f"{where}: agent_model_type must be 'excel' for this pipeline, "
            f"got {identity.agent_model_type!r}"
        )
    if identity.provider == "claude_excel_agent":
        if not identity.ui_model_label:
            raise AgentIdentityError(
                f"{where}: claude_excel_agent identities must pin "
                f"ui_model_label (the add-in model dropdown text)"
            )
        if identity.thinking_effort is not None:
            raise AgentIdentityError(
                f"{where}: thinking_effort is a ChatGPT axis; must be null "
                f"for claude_excel_agent"
            )
    else:  # chatgpt_excel_agent
        if not identity.thinking_effort:
            raise AgentIdentityError(
                f"{where}: chatgpt_excel_agent identities must pin "
                f"thinking_effort (the add-in's 'Thinking effort' label)"
            )
        # ui_model_label became a ChatGPT axis on 2026-08-27, when the
        # add-in grew a combined "Model and thinking effort" menu. Entries
        # from before then pin null; they stay loadable (the registry is
        # append-only history) but resolve_agent_identity refuses to run
        # them — without a pinned model the run would use whatever the
        # panel defaults to.
    return identity


def load_registry(path: Path = REGISTRY_PATH) -> dict[str, AgentIdentity]:
    """{agent_model_name: AgentIdentity}, refusing duplicate labels or axis
    combos (two labels for the same settings would split one cohort)."""
    entries = yaml.safe_load(path.read_text()) or []
    if not isinstance(entries, list):
        raise AgentIdentityError(f"{path}: expected a top-level list of entries")

    by_label: dict[str, AgentIdentity] = {}
    by_axes: dict[AxisKey, str] = {}
    for i, entry in enumerate(entries):
        where = f"{path.name} entry {i + 1}"
        if not isinstance(entry, dict):
            raise AgentIdentityError(f"{where}: expected a mapping")
        identity = _entry_to_identity(entry, where)
        if identity.agent_model_name in by_label:
            raise AgentIdentityError(
                f"{where}: agent_model_name '{identity.agent_model_name}' "
                f"already used — labels must be unique"
            )
        if identity.axes in by_axes:
            raise AgentIdentityError(
                f"{where}: axes {identity.axes} already mapped to "
                f"'{by_axes[identity.axes]}' — one combination, one cohort"
            )
        by_label[identity.agent_model_name] = identity
        by_axes[identity.axes] = identity.agent_model_name
    return by_label


def _stanza(label: str) -> str:
    return "\n".join([
        f"- agent_model_name: {label}",
        "  provider: claude_excel_agent      # or chatgpt_excel_agent",
        '  ui_model_label: "Opus 4.6"        # exact model dropdown/menu text (both providers)',
        "  thinking_effort: null             # ChatGPT 'Thinking effort' label; null for Claude",
        f"  agent_folder: {label}",
        "  agent_model_type: excel",
    ])


def _cfg_get(cfg: Any, key: str, default: Any = None) -> Any:
    if isinstance(cfg, dict):
        return cfg.get(key, default)
    return getattr(cfg, key, default)


def resolve_agent_identity(
    cfg: SimpleNamespace | dict[str, Any], path: Path = REGISTRY_PATH
) -> AgentIdentity:
    """The identity named by cfg.agent_model_name, or a refusal.

    Refuses when the label is unregistered (prints the stanza to add) or
    when the config sets any pinned key — including `model` /
    `thinking_effort` inside a provider block — since the registry is the
    only source for those.
    """
    label = _cfg_get(cfg, "agent_model_name")
    if not label:
        raise AgentIdentityError(
            f"Missing required config key: agent_model_name — the cohort "
            f"label from {path.name}. Registered: {sorted(load_registry(path))}"
        )

    present = [k for k in PINNED_KEY_NAMES if _cfg_get(cfg, k) is not None]
    for block_name in VALID_PROVIDERS:
        block = _cfg_get(cfg, block_name)
        if block is None:
            continue
        for key in _PROVIDER_BLOCK_PINNED:
            if _cfg_get(block, key) is not None:
                present.append(f"{block_name}.{key}")
    if present:
        raise AgentIdentityError(
            f"{', '.join(present)} may not be set in a run config: every "
            f"model setting is pinned by the agent_model_name entry in "
            f"{path.name}, so all rows under {label!r} ran the same way. "
            f"Remove the key(s); to run with different values, register a "
            f"NEW identity with a new label."
        )

    registry = load_registry(path)
    identity = registry.get(str(label))
    if identity is None:
        raise AgentIdentityError(
            f"No agent identity registered for agent_model_name={label!r}. "
            f"Registered: {sorted(registry)}. To add it, append to {path} "
            f"(fill in every field — none may be omitted):\n\n{_stanza(str(label))}\n"
        )
    if identity.provider == "chatgpt_excel_agent" and identity.ui_model_label is None:
        raise AgentIdentityError(
            f"agent_model_name={label!r} predates the ChatGPT add-in's model "
            f"picker (2026-08-27 UI) and pins no ui_model_label — without it "
            f"the run would use whatever model the panel defaults to. The "
            f"registry is append-only: leave that entry as history and "
            f"register a NEW label pinning both axes, e.g.:\n\n"
            + "\n".join([
                f"- agent_model_name: {label}_<model>",
                "  provider: chatgpt_excel_agent",
                '  ui_model_label: "GPT-5.6 Sol"     # exact Model menu text',
                f'  thinking_effort: "{identity.thinking_effort}"',
                f"  agent_folder: {label}_<model>",
                "  agent_model_type: excel",
            ])
            + "\n"
        )
    return identity
