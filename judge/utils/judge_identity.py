"""Resolve a grader's settings from its label (gradings.grader_model) via
the judge_identities.yaml registry.

`--model <label>` names the grader; the registry entry for that label pins
every setting that changes which endpoint is hit and how:

    provider, model (wire id), effort

The same label therefore always grades the same way — never a function of
which env vars happen to be exported (the old prefix-sniffing routing let
one slug hit different endpoints per shell). An unknown label refuses to
run and prints a paste-ready stanza.

Mirrors coding-agents/coding_agent/agent_identity.py.
"""

from pathlib import Path
from typing import Any, Dict, NamedTuple, Optional, Tuple

import yaml

REGISTRY_PATH = Path(__file__).resolve().parents[1] / "judge_identities.yaml"

# provider -> (base_url; None = the SDK default, i.e. api.openai.com,
#              api-key entry in repo_config.API_KEYS)
PROVIDERS = {
    "openrouter": ("https://openrouter.ai/api/v1", "openrouter"),
    "gemini": ("https://generativelanguage.googleapis.com/v1beta/openai/", "gemini"),
    "anthropic": ("https://api.anthropic.com/v1/", "anthropic"),
    "openai": (None, "openai"),
    # TensorBlock Forge — OpenAI-compatible gateway, separate key and
    # billing (Forge credits). Probed 2026-09-15 on tensorblock/gpt-5.6-sol:
    # reasoning_effort is forwarded (an invalid level is rejected upstream),
    # function tools work, usage carries cached_tokens, no per-call cost.
    # The upstream is OpenAI's chat/completions, so the judge treats it like
    # provider "openai" for request shaping (image-part stripping, tools
    # forcing effort "none") and routes real reasoning tiers through
    # /v1/responses like provider "openai" (Forge forwards encrypted
    # reasoning items; probed with a two-round tool exchange at high).
    "tensorblock": ("https://api.forge.tensorblock.co/v1", "forge"),
}

# (provider, model, effort) — the axes that make two graders different.
AxisKey = Tuple[str, str, Optional[str]]

# Every setting an identity pins, in stanza order. (key, type, nullable).
PINNED_KEYS = (
    ("provider", str, False),
    ("model", str, False),
    ("effort", str, True),
)
PINNED_KEY_NAMES = tuple(k for k, _, _ in PINNED_KEYS)


class JudgeIdentity(NamedTuple):
    grader_model: str
    provider: str
    model: str
    effort: Optional[str]

    @property
    def axes(self) -> AxisKey:
        return (self.provider, self.model, self.effort)

    @property
    def base_url(self) -> Optional[str]:
        return PROVIDERS[self.provider][0]

    @property
    def api_key_provider(self) -> str:
        return PROVIDERS[self.provider][1]

    def settings(self) -> Dict[str, Any]:
        """The pinned keys — recorded in _metadata.json / run summaries."""
        return {k: getattr(self, k) for k in PINNED_KEY_NAMES}


class JudgeIdentityError(ValueError):
    """Registry is invalid or the label is unknown."""


def _entry_to_identity(entry: Dict[str, Any], where: str) -> JudgeIdentity:
    label = entry.get("grader_model")
    if not label:
        raise JudgeIdentityError(f"{where}: entry is missing 'grader_model'")
    values: Dict[str, Any] = {}
    for key, typ, nullable in PINNED_KEYS:
        if key not in entry:
            hint = " (use null to not send it)" if nullable else ""
            raise JudgeIdentityError(
                f"{where}: entry is missing '{key}' — it changes what the "
                f"judge does, so every identity must pin it{hint}"
            )
        value = entry[key]
        if value is None:
            if not nullable:
                raise JudgeIdentityError(f"{where}: '{key}' may not be null")
        elif not isinstance(value, typ):
            raise JudgeIdentityError(
                f"{where}: '{key}' must be {typ.__name__}, got {value!r}"
            )
        values[key] = value
    if values["provider"] not in PROVIDERS:
        raise JudgeIdentityError(
            f"{where}: provider must be one of {sorted(PROVIDERS)}, "
            f"got {values['provider']!r}"
        )
    return JudgeIdentity(str(label), **values)


def load_registry(path: Path = REGISTRY_PATH) -> Dict[str, JudgeIdentity]:
    """{grader_model: JudgeIdentity}, refusing duplicate labels or axis
    combos (two labels for the same grader settings would split one grader
    population in the gradings table)."""
    entries = yaml.safe_load(path.read_text()) or []
    if not isinstance(entries, list):
        raise JudgeIdentityError(f"{path}: expected a top-level list of entries")

    by_label: Dict[str, JudgeIdentity] = {}
    by_axes: Dict[AxisKey, str] = {}
    for i, entry in enumerate(entries):
        where = f"{path.name} entry {i + 1}"
        if not isinstance(entry, dict):
            raise JudgeIdentityError(f"{where}: expected a mapping")
        identity = _entry_to_identity(entry, where)
        if identity.grader_model in by_label:
            raise JudgeIdentityError(
                f"{where}: grader_model '{identity.grader_model}' already "
                f"used — labels must be unique"
            )
        if identity.axes in by_axes:
            raise JudgeIdentityError(
                f"{where}: axes {identity.axes} already mapped to "
                f"'{by_axes[identity.axes]}' — one combination, one label"
            )
        by_label[identity.grader_model] = identity
        by_axes[identity.axes] = identity.grader_model
    return by_label


def _stanza(label: str) -> str:
    return "\n".join([
        f"- grader_model: {label}",
        "  provider: openrouter              # openrouter | gemini | anthropic | openai | tensorblock",
        "  model: <id sent on the wire>      # full slug for openrouter; bare id otherwise",
        "  effort: minimal                   # reasoning_effort; null = don't send",
    ])


def resolve_judge_identity(label: str, path: Path = REGISTRY_PATH) -> JudgeIdentity:
    """The identity for a grader label, or a refusal.

    Refuses when the label is unregistered (prints the stanza to add) so an
    unlisted `--model` fails fast instead of silently routing to OpenRouter.
    """
    if not label:
        raise JudgeIdentityError(
            f"No grader label given — pass --model with a label from "
            f"{path.name}. Registered: {sorted(load_registry(path))}"
        )
    registry = load_registry(path)
    identity = registry.get(str(label))
    if identity is None:
        raise JudgeIdentityError(
            f"No judge identity registered for grader_model={label!r}. "
            f"Registered: {sorted(registry)}. To add it, append to {path} "
            f"(fill in every field — none may be omitted):\n\n{_stanza(str(label))}\n"
        )
    return identity
