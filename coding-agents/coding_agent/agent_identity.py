"""Resolve a run's agent settings from its cohort label
(task_attempts.agent_model_name) via the agent_identities.yaml registry.

A run config names its cohort with `agent_model_name`; nothing else about
the agent may appear in the config. The registry entry for that label
supplies every setting that changes what the agent does:

    cli, model, effort, extra_args, env

The runner builds its AgentConfig from them, and the attempt row records the
same dict in its extra_configs. An unknown
label refuses to run and prints a paste-ready stanza; a config that sets an
`agent:` block refuses to run too, so two rows under one label can never
have run with different settings.

Mirrors cli-agents/excel_cli_agent/agent_identity.py.
"""

from pathlib import Path
from typing import Any, Dict, List, NamedTuple, Optional, Tuple

import yaml

REGISTRY_PATH = Path(__file__).resolve().parent / "agent_identities.yaml"

# (cli, model, effort) — the axes that make two cohorts different.
AxisKey = Tuple[str, str, Optional[str]]

# Every setting an identity pins, in stanza order. (key, type, nullable).
PINNED_KEYS = (
    ("cli", str, False),
    ("model", str, False),
    ("effort", str, True),
    ("extra_args", list, False),
    ("env", dict, False),
)
PINNED_KEY_NAMES = tuple(k for k, _, _ in PINNED_KEYS)

# Keys a run config may not carry: the identity is the only source for them.
FORBIDDEN_CONFIG_KEYS = ("agent", "identity") + PINNED_KEY_NAMES


class AgentIdentity(NamedTuple):
    agent_model_name: str
    cli: str
    model: str
    effort: Optional[str]
    extra_args: List[str]
    env: Dict[str, str]

    @property
    def axes(self) -> AxisKey:
        return (self.cli, self.model, self.effort)

    def settings(self) -> Dict[str, Any]:
        """The pinned keys, and the dict stored in task_attempts.extra_configs."""
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
            hint = " (use null if the CLI does not take it)" if nullable else ""
            raise AgentIdentityError(
                f"{where}: entry is missing '{key}' — it changes what the agent "
                f"does, so every identity must pin it{hint}"
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
    if values["cli"] not in ("claude", "codex"):
        raise AgentIdentityError(f"{where}: cli must be claude|codex, got {values['cli']!r}")
    values["extra_args"] = [str(a) for a in values["extra_args"]]
    values["env"] = {str(k): str(v) for k, v in values["env"].items()}
    return AgentIdentity(str(label), **values)


def load_registry(path: Path = REGISTRY_PATH) -> Dict[str, AgentIdentity]:
    """{agent_model_name: AgentIdentity}, refusing duplicate labels or axis
    combos (two labels for the same agent settings would split one cohort)."""
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
        "  cli: claude                       # or codex",
        "  model: <vendor model id>",
        "  effort: null                      # claude: low|medium|high|xhigh|max; codex: model_reasoning_effort",
        "  extra_args: []",
        "  env: {}",
    ])


def resolve_agent_identity(config: Dict[str, Any], path: Path = REGISTRY_PATH) -> AgentIdentity:
    """The identity named by config['agent_model_name'], or a refusal.

    Refuses when the label is unregistered (prints the stanza to add) or when
    the config sets an agent block / any pinned key — the registry is the only
    source for those, so a config cannot even repeat them.
    """
    label = config.get("agent_model_name")
    if not label:
        raise AgentIdentityError(
            f"Missing required field in config: agent_model_name — the cohort "
            f"label from {path.name}. Registered: "
            f"{sorted(load_registry(path))}"
        )
    present = [k for k in FORBIDDEN_CONFIG_KEYS if k in config]
    if present:
        raise AgentIdentityError(
            f"{', '.join(present)} may not be set in a run config: every "
            f"agent setting is pinned by the agent_model_name entry in "
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
