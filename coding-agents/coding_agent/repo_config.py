"""Read the repo config at <MBABench>/config/config.yaml and resolve the two
local roots every run uses.

coding-agents is a workspace member of MBABench, whose root pyproject.toml
exposes config/python/config.py as the top-level `config` module. That file
holds the model API keys (keys.*) and the two local roots:

  local.data_root    the downloaded benchmark (tasks/task_id=<N>/task.json ...)
  local.output_root  where runs write attempts/<agent_model_name>/...

This is a deliberate copy of cli-agents/excel_cli_agent/repo_config.py
(kept separate so each package stays runnable on its own).

RESOLUTION ORDER for the roots, first hit wins:

  1. the environment (MBABENCH_DATA_ROOT / MBABENCH_OUTPUT_ROOT)
  2. the repo config (local.data_root / local.output_root)
  3. the defaults "data" / "outputs"

A relative value resolves against the repo root. Model API keys resolve the
other way round: the environment (or .env) first, then config/config.yaml
keys.* — see config.resolve_api_key.
"""

import logging
import os
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Points the repo-config lookup at a different directory. An escape hatch for
# layouts where the installed location is wrong. Unset -> Config resolves its
# own directory via the editable install.
REPO_CONFIG_DIR_ENV = "MBABENCH_CONFIG_DIR"

# Per-run overrides for the two local roots (see the module docstring).
DATA_ROOT_ENV = "MBABENCH_DATA_ROOT"
OUTPUT_ROOT_ENV = "MBABENCH_OUTPUT_ROOT"

DEFAULT_DATA_ROOT = "data"
DEFAULT_OUTPUT_ROOT = "outputs"


def repo_value(*path: str) -> Optional[str]:
    """Non-empty string at a path in <repo>/config/config.yaml, else None.

    `from config import Config` works because the root pyproject.toml exposes
    config/python/config.py as a top-level module, so an editable install
    resolves it from any cwd. This wrapper adds the three things that import
    does not give you:

    * It never raises. A standalone checkout has no config/
      directory, so the import fails there — and that failure IS the signal
      to fall through to the environment, not an error.
    * It never writes. Config.load() defaults to create_missing=True, which
      seeds a config.yaml from the defaults as a side effect of reading one.
    * A `null` placeholder yields None, so an unset key falls through to the
      next layer rather than resolving to something falsy-but-present.
    """
    try:
        from config import Config
    except ImportError:
        logger.debug("repo `config` module not installed; using env vars")
        return None

    override = os.environ.get(REPO_CONFIG_DIR_ENV)
    # Config.load warns about every unset ${env:VAR} in the file, including
    # keys coding-agents never reads (gemini_api_key, ...). Not worth a
    # warning to someone starting a run.
    cfg_log = logging.getLogger("config")
    prev_level = cfg_log.level
    cfg_log.setLevel(logging.ERROR)
    try:
        data = Config.load(
            Path(override).expanduser() if override else None,
            create_missing=False,
            check_required=False,
        ).as_dict()
    except Exception as e:  # degrade, never break the caller
        logger.warning(
            "could not read the repo config (%s: %s); falling back to "
            "environment variables",
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


def monorepo_root() -> Optional[Path]:
    """<MBABench>, the directory the shared assets hang off (house_standards/
    ...), or None on a standalone checkout.

    The shared `config` module lives at <root>/config/python/config.py, so
    its DEFAULT_CONFIG_DIR is the authoritative locator — deliberately not
    MBABENCH_CONFIG_DIR, which redirects only the yaml lookup. Without the
    workspace install, a checkout sitting directly under the repo (this file
    lives at <root>/coding-agents/coding_agent/) still resolves through its
    parent.
    """
    try:
        from config import Config
    except ImportError:
        candidate = Path(__file__).resolve().parents[2]
        if (candidate / "config" / "config_default.yaml").exists():
            return candidate
        logger.debug("repo `config` module not installed and no repo above the checkout")
        return None
    return Path(Config.DEFAULT_CONFIG_DIR).resolve().parent


def _local_root(env_var: str, key: str, default: str) -> Path:
    """A `local.*` root as an absolute path: env var, then config, then the
    default. A relative value resolves against the repo root (and, on a
    standalone checkout with no root to resolve against, the cwd)."""
    raw = os.environ.get(env_var) or repo_value("local", key) or default
    path = Path(raw).expanduser()
    if path.is_absolute():
        return path
    root = monorepo_root()
    return (root / path) if root is not None else path.resolve()


def data_root() -> Path:
    """Where the downloaded benchmark lives (`local.data_root`)."""
    return _local_root(DATA_ROOT_ENV, "data_root", DEFAULT_DATA_ROOT)


def output_root() -> Path:
    """Where a run writes its attempts (`local.output_root`)."""
    return _local_root(OUTPUT_ROOT_ENV, "output_root", DEFAULT_OUTPUT_ROOT)


def describe_local_target() -> str:
    """One log line naming the two roots a run will use, and why."""
    src_in = (DATA_ROOT_ENV if os.environ.get(DATA_ROOT_ENV)
              else "config/config.yaml local.data_root" if repo_value("local", "data_root")
              else "default")
    src_out = (OUTPUT_ROOT_ENV if os.environ.get(OUTPUT_ROOT_ENV)
               else "config/config.yaml local.output_root" if repo_value("local", "output_root")
               else "default")
    return f"{data_root()} (from {src_in}) -> {output_root()} (from {src_out})"


def data_path(rel: str) -> Path:
    """Absolute path for a data-root-relative POSIX path recorded in task.json
    (`tasks/task_id=1/starting_files/Foo.xlsx`)."""
    p = Path(str(rel).replace("\\", "/"))
    return p if p.is_absolute() else data_root() / p


def output_relative(path: Path) -> str:
    """POSIX path relative to the output root, or an absolute one when the
    file sits outside it (an output_root pointed elsewhere)."""
    path = Path(path).resolve()
    try:
        return path.relative_to(output_root().resolve()).as_posix()
    except ValueError:
        return path.as_posix()
