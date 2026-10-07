"""Read the repo config at <MBABench>/config/config.yaml and resolve the two
local roots every run uses.

cli-agents is a workspace member of MBABench, whose root pyproject.toml
exposes config/python/config.py as the top-level `config` module. That file
holds the model API keys (keys.*) and the two local roots:

  local.data_root    the downloaded benchmark (tasks/task_id=<N>/task.json ...)
  local.output_root  where runs write attempts/<agent_model_name>/...

RESOLUTION ORDER for the roots, first hit wins:

  1. the environment (MBABENCH_DATA_ROOT / MBABENCH_OUTPUT_ROOT) — a one-run
     override, e.g. a benchmark unpacked on an external disk.
  2. the repo config (local.data_root / local.output_root).
  3. the defaults "data" / "outputs".

A relative value resolves against the repo root, so the same config works
from any cwd. Model API keys resolve the other way round: the environment
(or .env) first, then config/config.yaml keys.*.
"""

import hashlib
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

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
    # keys cli-agents never reads (gemini_api_key, ...). Not worth a warning
    # to someone starting a batch run.
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


def monorepo_root() -> Path:
    """<MBABench>, the directory holding config/ and house_standards/.

    Asks the installed `config` module for its directory (config/python/
    config.py -> <root>/config), so the root can't disagree with the config
    the keys come from. A standalone checkout has no `config` module; there
    the workspace layout (this file lives at <root>/cli-agents/
    excel_cli_agent/) is the only answer.
    """
    try:
        from config import Config
        return Path(Config.DEFAULT_CONFIG_DIR).resolve().parent
    except (ImportError, AttributeError):
        return Path(__file__).resolve().parents[2]


def _resolve_root(env_name: str, config_key: str, default: str) -> Path:
    """One of the two local roots, absolute. See the module docstring.

    An absolute value is taken as-is; a relative one hangs off the repo
    root, so `data` means the same directory whether the batch was launched
    from cli-agents/ or from the repo root.
    """
    raw = os.environ.get(env_name) or repo_value("local", config_key) or default
    path = Path(raw).expanduser()
    return path if path.is_absolute() else (monorepo_root() / path)


def data_root() -> Path:
    """Where the downloaded benchmark lives (tasks/task_id=<N>/...)."""
    return _resolve_root(DATA_ROOT_ENV, "data_root", DEFAULT_DATA_ROOT)


def output_root() -> Path:
    """Where a run writes its attempts (attempts/<agent_model_name>/...)."""
    return _resolve_root(OUTPUT_ROOT_ENV, "output_root", DEFAULT_OUTPUT_ROOT)


def describe_local_target() -> str:
    """One log line naming the two roots a run will use, and why — so an
    empty outputs/ never looks like a run that did nothing."""
    src_in = (DATA_ROOT_ENV if os.environ.get(DATA_ROOT_ENV)
              else "config/config.yaml local.data_root" if repo_value("local", "data_root")
              else "default")
    src_out = (OUTPUT_ROOT_ENV if os.environ.get(OUTPUT_ROOT_ENV)
               else "config/config.yaml local.output_root" if repo_value("local", "output_root")
               else "default")
    return f"{data_root()} (from {src_in}) -> {output_root()} (from {src_out})"


def resolve_data_path(rel_path: str) -> Path:
    """Absolute path for a data-root-relative POSIX path out of task.json
    (`tasks/task_id=1/starting_files/Foo.xlsx`)."""
    rel = Path(str(rel_path).replace("\\", "/"))
    return rel if rel.is_absolute() else data_root() / rel


def to_output_relative(path: Path) -> str:
    """POSIX spelling of a path a run wrote, relative to the output root.

    Every attempt_files / prompt_files entry uses this form, so the output
    tree stays valid when it moves. A path outside the output root has no
    relative spelling and is kept absolute.
    """
    path = Path(path).resolve()
    try:
        return path.relative_to(output_root().resolve()).as_posix()
    except ValueError:
        return path.as_posix()


def resolve_attachments(rel_paths: Iterable[str]) -> List[Path]:
    """Absolute paths for repo-root-relative attachment paths.

    Raises at once on a missing or empty file: an attachment is prompt text
    the recorded prompt_version promises the agent saw, so running without
    it would be the empty-context defect all over again — refuse before any
    task is claimed.
    """
    root = monorepo_root()
    out: List[Path] = []
    for rel in rel_paths:
        path = root / rel
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(
                f"Prompt attachment {rel} is missing or empty under {root} "
                "(the prompt version declares it; it must exist before the "
                "batch starts)"
            )
        out.append(path)
    return out


def attachment_extra_configs(paths: Iterable[Path],
                             names: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    """The provenance keys merged into an attempt row's extra_configs.

    House_Standards_v<n>.md -> house_standards: {version, file, sha256},
    the record every pipeline writes. `file` is the name the agent saw in
    the workspace (v15+ delivers it as HOUSE_STANDARDS.md; `source` then
    keeps the versioned source name). The hash is computed at run time from
    the file actually shipped, not copied from a constant, so a silently
    edited file shows up as a different sha.
    """
    names = names or {}
    out: Dict[str, Any] = {}
    for path in paths:
        m = re.fullmatch(r"House_Standards_v(\d+)\.md", path.name)
        if not m:
            continue
        delivered = names.get(path.name, path.name)
        rec = {
            "version": int(m.group(1)),
            "file": delivered,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        }
        if delivered != path.name:
            rec["source"] = path.name
        out["house_standards"] = rec
    return out
