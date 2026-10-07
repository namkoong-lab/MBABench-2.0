"""Derive the agent identity from the behavior-determining fields in cfg.

The identity is a pure function of the config fields that change agent
output, so `task_attempts.agent_model_name` cannot disagree with what the
run actually did. Labels bifurcate on the UI axes that change agent output:
Claude mode (chat/cowork) + model + effort; ChatGPT mode (chat/work) + model
+ intelligence (chat) or effort + speed (work).

The tables below are APPEND-ONLY. Every label is referenced by attempt rows
already recorded, so editing one silently rewrites what those rows mean. To
add a cohort, add an entry. Unknown combinations raise
`UnknownAgentCombination`, which forces a naming decision before an
unclassified label is written.
"""

from __future__ import annotations

from dataclasses import dataclass
from types import SimpleNamespace


@dataclass(frozen=True)
class AgentIdentity:
    model_name: str  # → task_attempts.agent_model_name
    agent_folder: str  # → the engine's session.agent_name (stamped into filenames)
    agent_model_type: str = "gui"  # → task_attempts.agent_model_type


class UnknownAgentCombination(ValueError):
    pass


# Signature: (claude_web.mode, claude_web.model, claude_web.effort). Mode is
# in the key so chat and cowork cohorts stay separable; the chat labels carry
# no mode segment.
#
# Effort is a reasoning axis that changes agent output, so a max-effort answer
# is not the same result as a medium-effort one and the two must not share a
# label. Haiku is the exception: claude.ai exposes no Effort control for it,
# so the config value is inert there and stays out of the label.
_CLAUDE_IDENTITIES: dict[tuple, AgentIdentity] = {
    ("chat", "sonnet_4_6", "max"): AgentIdentity(
        "claude_sonnet_4_6_max", "claude_sonnet_4_6_max"
    ),
    ("chat", "opus_4_6", "max"): AgentIdentity(
        "claude_opus_4_6_max", "claude_opus_4_6_max"
    ),
    ("chat", "opus_4_8", "max"): AgentIdentity(
        "claude_opus_4_8_max", "claude_opus_4_8_max"
    ),
    ("chat", "haiku_4_5", "max"): AgentIdentity(
        "claude_haiku_4_5", "claude_haiku_4_5"
    ),
    ("chat", "haiku_4_5", None): AgentIdentity(
        "claude_haiku_4_5", "claude_haiku_4_5"
    ),
    ("chat", "fable_5", "max"): AgentIdentity(
        "claude_fable_5_max", "claude_fable_5_max"
    ),
    ("cowork", "fable_5", "max"): AgentIdentity(
        "claude_fable_5_cowork_max", "claude_fable_5_cowork_max"
    ),
    ("cowork", "opus_5", "max"): AgentIdentity(
        "claude_opus_5_cowork_max", "claude_opus_5_cowork_max"
    ),
    # 2026-09-10 (101-task rerun): Fable 5.1 in cowork at Max. Probed live
    # on claude.ai the same day — top-level radio "Fable 5.1", Effort flyout
    # Low/Medium/High/Extra/Max, button reads "Model: Fable 5.1 Max". Chat
    # mode deliberately has no entry: the rerun is cowork-only.
    ("cowork", "fable_5_1", "max"): AgentIdentity(
        "claude_fable_5_1_cowork_max", "claude_fable_5_1_cowork_max"
    ),
    # 2026-09-23: Opus 5.5 in cowork at Max. Cowork-only, like the Fable 5.1
    # rerun above — a chat-mode config naming opus_5_5 is refused before the
    # browser opens.
    ("cowork", "opus_5_5", "max"): AgentIdentity(
        "claude_opus_5_5_cowork_max", "claude_opus_5_5_cowork_max"
    ),
}


# Signatures (mode defaults to "chat"):
#   chat: (chatgpt_web.mode, chatgpt_web.model, chatgpt_web.intelligence)
#   work: (chatgpt_web.mode, chatgpt_web.model, chatgpt_web.effort,
#          chatgpt_web.speed)
# model=None means "let the session default win". Intelligence is the
# chat-mode reasoning axis and changes agent output, so it bifurcates the
# label; intelligence=None keeps the bare label for the cohorts that pinned
# no intelligence. Work mode's axes are effort + speed; speed stays out of
# the label while it is "standard" (the UI default).
_CHATGPT_IDENTITIES: dict[tuple, AgentIdentity] = {
    ("chat", None, None): AgentIdentity("chatgpt_web", "chatgpt_web"),
    ("chat", "instant", None): AgentIdentity(
        "chatgpt_instant", "chatgpt_instant"
    ),
    ("chat", "thinking", None): AgentIdentity(
        "chatgpt_thinking", "chatgpt_thinking"
    ),
    ("chat", "pro", None): AgentIdentity("chatgpt_web_pro", "chatgpt_web_pro"),
    ("chat", "gpt_5_6_sol", None): AgentIdentity(
        "chatgpt_gpt_5_6_sol", "chatgpt_gpt_5_6_sol"
    ),
    # Sol with the chat tier actually pinned. Distinct from the entry above,
    # which is the cohort that let the session default stand.
    ("chat", "gpt_5_6_sol", "pro"): AgentIdentity(
        "chatgpt_gpt_5_6_sol_pro", "chatgpt_gpt_5_6_sol_pro"
    ),
    ("chat", "gpt_5_5", None): AgentIdentity(
        "chatgpt_gpt_5_5", "chatgpt_gpt_5_5"
    ),
    ("chat", "gpt_5_5", "instant"): AgentIdentity(
        "chatgpt_gpt_5_5_instant", "chatgpt_gpt_5_5_instant"
    ),
    ("work", "gpt_5_6_sol", "ultra", "standard"): AgentIdentity(
        "chatgpt_gpt_5_6_sol_work_ultra", "chatgpt_gpt_5_6_sol_work_ultra"
    ),
    # 2026-09-10 (101-task rerun): GPT-6 Astra in work mode at Ultra.
    # Probed live on chatgpt.com the same day — slider picker, model radio
    # "GPT-6 Astra", Power ladder Light..Ultra (6 stops), pill reads
    # "GPT-6 Astra Ultra" and held Ultra across a menu reopen.
    ("work", "gpt_6_astra", "ultra", "standard"): AgentIdentity(
        "chatgpt_gpt_6_astra_work_ultra", "chatgpt_gpt_6_astra_work_ultra"
    ),
    # 2026-09-21 (101-task chat cohort): the GPT-6 generation in CHAT mode at
    # Pro. Chat has the slider picker; its model list is Latest / GPT-5.6 Sol /
    # GPT-5.5, the ladder under Latest is Instant..Pro (5 stops), and the
    # pill at the top stop reads "6Pro". Work mode's "GPT-6 Astra" radio does
    # not exist in chat, so this label says gpt_6, not gpt_6_astra: it claims
    # only what the UI shows. The agent refuses to send unless the pill reads
    # exactly "6Pro".
    ("chat", "gpt_6", "pro"): AgentIdentity(
        "chatgpt_gpt_6_pro", "chatgpt_gpt_6_pro"
    ),
}


def resolve_agent_identity(cfg: SimpleNamespace) -> AgentIdentity:
    provider = getattr(getattr(cfg, "provider", None), "kind", None)
    if provider == "claude":
        return _resolve_claude(cfg)
    if provider == "chatgpt":
        return _resolve_chatgpt(cfg)
    raise UnknownAgentCombination(
        f"provider.kind={provider!r} has no identity resolver. "
        f"Add one in infra/configs/agent_identity.py."
    )


def _claude_block(cfg: SimpleNamespace) -> SimpleNamespace:
    block = getattr(cfg, "claude_web", None)
    if block is None:
        raise UnknownAgentCombination(
            "provider=claude but cfg.claude_web block is missing."
        )
    return block


def _chatgpt_block(cfg: SimpleNamespace) -> SimpleNamespace:
    block = getattr(cfg, "chatgpt_web", None)
    if block is None:
        raise UnknownAgentCombination(
            "provider=chatgpt but cfg.chatgpt_web block is missing."
        )
    return block


def _resolve_claude(cfg: SimpleNamespace) -> AgentIdentity:
    block = _claude_block(cfg)
    model = getattr(block, "model", None)
    mode = (getattr(block, "mode", None) or "chat").lower()
    effort = getattr(block, "effort", None)
    key = (mode, model, effort)
    try:
        return _CLAUDE_IDENTITIES[key]
    except KeyError:
        raise UnknownAgentCombination(
            f"No Claude identity for (claude_web.mode, claude_web.model, "
            f"claude_web.effort)={key!r}. "
            f"Known: {list(_CLAUDE_IDENTITIES)}. "
            f"Add an entry in infra/configs/agent_identity.py "
            f"if this is a real combination."
        )


def _resolve_chatgpt(cfg: SimpleNamespace) -> AgentIdentity:
    block = _chatgpt_block(cfg)
    mode = (getattr(block, "mode", None) or "chat").lower()
    model = getattr(block, "model", None)
    if mode == "work":
        effort = getattr(block, "effort", None)
        # null speed is the UI's "standard", so both spell the same cohort.
        speed = getattr(block, "speed", None) or "standard"
        key = (mode, model, effort, speed)
        axes = "(mode, model, effort, speed)"
    else:
        key = (mode, model, getattr(block, "intelligence", None))
        axes = "(mode, model, intelligence)"
    try:
        return _CHATGPT_IDENTITIES[key]
    except KeyError:
        raise UnknownAgentCombination(
            f"No ChatGPT identity for chatgpt_web {axes}={key!r}. "
            f"Known: {list(_CHATGPT_IDENTITIES)}. "
            f"Add an entry in infra/configs/agent_identity.py "
            f"if this is a real combination."
        )
