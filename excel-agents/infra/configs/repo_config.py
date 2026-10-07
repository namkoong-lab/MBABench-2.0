"""Where the benchmark data lives and where runs write, from the shared config.

The repo-level config (<repo>/config/config.yaml, exposed as the top-level
`config` module by the root pyproject.toml) names both roots under `local:`.
Resolution, first hit wins:

    data_root    $MBABENCH_DATA_ROOT   -> local.data_root   -> <repo>/data
    output_root  $MBABENCH_OUTPUT_ROOT -> local.output_root -> <repo>/outputs

A relative value resolves against the repository root (the directory holding
config/ and house_standards/), so the same setting works from any cwd.
`MBABENCH_CONFIG_DIR` points the config lookup at another directory.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)

REPO_CONFIG_DIR_ENV = "MBABENCH_CONFIG_DIR"
DATA_ROOT_ENV = "MBABENCH_DATA_ROOT"
OUTPUT_ROOT_ENV = "MBABENCH_OUTPUT_ROOT"
DEFAULT_DATA_ROOT = "data"
DEFAULT_OUTPUT_ROOT = "outputs"

# <repo>/excel-agents/infra/configs/repo_config.py -> <repo>
_FALLBACK_REPO_ROOT = Path(__file__).resolve().parents[3]


def repo_root() -> Path:
    """The repository root: the directory holding config/ and house_standards/.

    Asks the installed `config` module for its directory so the answer cannot
    disagree with the config it reads; without that module (a bare checkout)
    the workspace layout is the only answer.
    """
    override = os.environ.get(REPO_CONFIG_DIR_ENV)
    if override:
        return Path(override).expanduser().resolve().parent
    try:
        from config import Config

        return Path(Config.DEFAULT_CONFIG_DIR).resolve().parent
    except (ImportError, AttributeError):
        return _FALLBACK_REPO_ROOT


def repo_value(*path: str) -> str | None:
    """Non-empty string at a dotted path in <repo>/config/config.yaml, else None.

    Never raises (a missing `config` module or an unreadable file falls
    through to the defaults) and never writes (Config.load() otherwise seeds
    a config.yaml as a side effect of reading one).
    """
    try:
        from config import Config
    except ImportError:
        logger.debug("shared `config` module not installed; using defaults")
        return None

    override = os.environ.get(REPO_CONFIG_DIR_ENV)
    # Config.load warns about every unset ${env:VAR} in the file, including
    # keys this pipeline never reads. Not worth a warning per run.
    cfg_log = logging.getLogger("config")
    prev_level = cfg_log.level
    cfg_log.setLevel(logging.ERROR)
    try:
        data = Config.load(
            Path(override).expanduser() if override else None,
            create_missing=False,
            check_required=False,
        ).as_dict()
    except Exception as e:  # noqa: BLE001 — degrade, never break the caller
        logger.warning(
            "could not read the shared config (%s: %s); using defaults",
            type(e).__name__,
            e,
        )
        return None
    finally:
        cfg_log.setLevel(prev_level)

    node = data
    for key in path:
        if not isinstance(node, dict):
            return None
        node = node.get(key)
    return node.strip() if isinstance(node, str) and node.strip() else None


def _resolve_root(env_name: str, config_key: str, default: str) -> tuple[Path, str]:
    raw = os.environ.get(env_name, "").strip()
    if raw:
        where = f"${env_name}"
    else:
        raw = repo_value("local", config_key)
        if raw:
            where = f"config/config.yaml local.{config_key}"
        else:
            raw, where = default, "default"
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = repo_root() / path
    return path.resolve(), where


def data_root() -> Path:
    """<data_root>: holds tasks/task_id=<N>/ (scripts/download_dataset.py)."""
    return _resolve_root(DATA_ROOT_ENV, "data_root", DEFAULT_DATA_ROOT)[0]


def output_root() -> Path:
    """<output_root>: attempts/<agent_model_name>/... are written under it."""
    return _resolve_root(OUTPUT_ROOT_ENV, "output_root", DEFAULT_OUTPUT_ROOT)[0]


def describe_local_roots() -> str:
    """One log line naming both roots and where each came from."""
    din, src_in = _resolve_root(DATA_ROOT_ENV, "data_root", DEFAULT_DATA_ROOT)
    dout, src_out = _resolve_root(OUTPUT_ROOT_ENV, "output_root", DEFAULT_OUTPUT_ROOT)
    return f"tasks from {din} ({src_in}); attempts under {dout} ({src_out})"
