"""Build the TaskSource / AttemptSink for a run from the loaded cfg namespace.

There is one source and one sink:

    source.kind: bundle   tasks from <data_root>/tasks/task_id=<N>/task.json
    sink.kind:   local    attempts under <output_root>/attempts/<agent_model_name>/

Both roots come from the shared config (infra/configs/repo_config.py):
$MBABENCH_DATA_ROOT / $MBABENCH_OUTPUT_ROOT, then local.data_root /
local.output_root in <repo>/config/config.yaml, then <repo>/data and
<repo>/outputs. Both builders take the full `cfg` so they can resolve the
agent identity (the cohort label the rows are written under).
"""

from __future__ import annotations

import logging
from types import SimpleNamespace

from infra.configs import data_root, output_root, resolve_agent_identity

from .base import AttemptSink, TaskSource

logger = logging.getLogger(__name__)

SOURCE_KINDS = ("bundle",)
SINK_KINDS = ("local",)


def _kind(cfg: SimpleNamespace, slot: str, valid: tuple[str, ...]) -> str:
    block = getattr(cfg, slot, None)
    kind = getattr(block, "kind", None) if block is not None else None
    if kind not in valid:
        raise ValueError(
            f"Unknown {slot}.kind: {kind!r}. Available: {list(valid)} — this "
            f"public release reads tasks from the downloaded benchmark and "
            f"writes attempts to local files only."
        )
    return kind


def build_source(cfg: SimpleNamespace) -> TaskSource:
    _kind(cfg, "source", SOURCE_KINDS)
    from .sources.bundle_source import BundleTaskSource

    filters = getattr(cfg.source, "filters", None) or SimpleNamespace()
    identity = resolve_agent_identity(cfg)
    return BundleTaskSource(
        data_root=data_root(),
        output_root=output_root(),
        agent_model_name=identity.model_name,
        prompt_version=cfg.agent.prompt_version,
        task_ids=list(getattr(filters, "task_ids", []) or []),
        skip_already_attempted=bool(getattr(filters, "skip_already_attempted", True)),
    )


def build_sink(cfg: SimpleNamespace) -> AttemptSink:
    _kind(cfg, "sink", SINK_KINDS)
    from .sinks.local_sink import LocalAttemptSink

    identity = resolve_agent_identity(cfg)
    return LocalAttemptSink(
        output_root=output_root(),
        agent_model_name=identity.model_name,
        agent_model_type=identity.agent_model_type,
        prompt_version=cfg.agent.prompt_version,
        # The identity's pinned settings (provider, ui_model_label,
        # thinking_effort, ...) are recorded in every row's extra_configs.
        extra_configs=identity.settings(),
    )
