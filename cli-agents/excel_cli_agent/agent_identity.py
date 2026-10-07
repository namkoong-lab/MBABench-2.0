"""Resolve a run's full model settings from its cohort label
(task_attempts.agent_model_name) via the agent_identities.yaml registry.

A batch config names its cohort with `agent_model_name`; nothing else about
the model may appear in the config. The registry entry for that label
supplies every setting that changes what the agent does:

    model, reasoning_effort, thinking_budget_tokens, max_completion_tokens,
    base_url, fresh_context_mode, enhanced_excel_context, recent_history_count

The runner copies them into the config, the executor reads them like any
other key, and the DB row records the same dict in
task_attempts.extra_configs. An unknown label refuses to run and prints a
paste-ready stanza; a config that sets any pinned key refuses to run too, so
two rows under one label can never have run with different settings.
"""

from pathlib import Path
from typing import Any, Dict, NamedTuple, Optional, Tuple

import yaml

REGISTRY_PATH = Path(__file__).resolve().parent / "agent_identities.yaml"

# (model, reasoning_effort, thinking_budget_tokens, max_completion_tokens)
AxisKey = Tuple[str, Optional[str], Optional[int], int]

# Every setting an identity pins, in stanza order. (key, type, nullable).
# A batch config must not set any of these.
PINNED_KEYS = (
    ("model", str, False),
    ("reasoning_effort", str, True),
    ("thinking_budget_tokens", int, True),
    ("max_completion_tokens", int, False),
    ("base_url", str, False),
    ("fresh_context_mode", bool, False),
    ("enhanced_excel_context", bool, False),
    ("recent_history_count", int, False),
)
PINNED_KEY_NAMES = tuple(k for k, _, _ in PINNED_KEYS)


class AgentIdentity(NamedTuple):
    agent_model_name: str
    model: str
    reasoning_effort: Optional[str]
    thinking_budget_tokens: Optional[int]
    max_completion_tokens: int
    base_url: str
    fresh_context_mode: bool
    enhanced_excel_context: bool
    recent_history_count: int

    @property
    def axes(self) -> AxisKey:
        return (self.model, self.reasoning_effort,
                self.thinking_budget_tokens, self.max_completion_tokens)

    def settings(self) -> Dict[str, Any]:
        """The pinned keys the runner writes into the batch config, and the
        dict stored in task_attempts.extra_configs."""
        return {k: getattr(self, k) for k in PINNED_KEY_NAMES}

    # Alias so call sites read naturally at the DB write.
    extra_configs = settings


class AgentIdentityError(ValueError):
    """Registry is invalid, the label is unknown, or the config sets a pinned key."""


def _entry_to_identity(entry: Dict[str, Any], where: str) -> AgentIdentity:
    label = entry.get("agent_model_name")
    if not label:
        raise AgentIdentityError(f"{where}: entry is missing 'agent_model_name'")
    values: Dict[str, Any] = {}
    for key, typ, nullable in PINNED_KEYS:
        if key not in entry:
            raise AgentIdentityError(
                f"{where}: entry is missing '{key}' — it changes what the agent "
                f"does, so every identity must pin it (use null if the model "
                f"does not take it)" if nullable else
                f"{where}: entry is missing '{key}' — it changes what the agent "
                f"does, so every identity must pin it"
            )
        value = entry[key]
        if value is None:
            if not nullable:
                raise AgentIdentityError(f"{where}: '{key}' may not be null")
        elif (typ is int and isinstance(value, bool)) or not isinstance(value, typ):
            # bool is an int subclass; keep the two strictly apart.
            raise AgentIdentityError(
                f"{where}: '{key}' must be {typ.__name__}, got {value!r}"
            )
        values[key] = value
    return AgentIdentity(str(label), **values)


def load_registry(path: Path = REGISTRY_PATH) -> Dict[str, AgentIdentity]:
    """{agent_model_name: AgentIdentity}, refusing duplicate labels or axis
    combos (two labels for the same model settings would split one cohort)."""
    entries = yaml.safe_load(path.read_text()) or []
    if not isinstance(entries, list):
        raise AgentIdentityError(f"{path}: expected a top-level list of entries")

    by_label: Dict[str, AgentIdentity] = {}
    by_axes: Dict[AxisKey, str] = {}
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
        "  model: <provider model id>",
        "  reasoning_effort: null            # or none/low/medium/high/xhigh/max",
        "  thinking_budget_tokens: null      # or an int < max_completion_tokens",
        "  max_completion_tokens: 64000",
        "  base_url: https://<provider endpoint>",
        "  fresh_context_mode: true",
        "  enhanced_excel_context: true",
        "  recent_history_count: 3",
    ])


def resolve_agent_identity(config: Dict[str, Any], path: Path = REGISTRY_PATH) -> AgentIdentity:
    """The identity named by config['agent_model_name'], or a refusal.

    Refuses when the label is unregistered (prints the stanza to add) or when
    the config sets any pinned key — the registry is the only source for
    those, so a config cannot even repeat them.
    """
    label = config.get("agent_model_name")
    if not label:
        raise AgentIdentityError(
            f"Missing required field in config: agent_model_name — the cohort "
            f"label from {path.name}. Registered: "
            f"{sorted(load_registry(path))}"
        )
    present = [k for k in PINNED_KEY_NAMES if k in config]
    if present:
        raise AgentIdentityError(
            f"{', '.join(present)} may not be set in a batch config: every "
            f"model setting is pinned by the agent_model_name entry in "
            f"{path.name}, so all rows under '{label}' ran the same way. "
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
    return identity


def resolve_agent_model_name(config: Dict[str, Any], path: Path = REGISTRY_PATH) -> str:
    """The cohort label, validated against the registry; see resolve_agent_identity."""
    return resolve_agent_identity(config, path).agent_model_name
