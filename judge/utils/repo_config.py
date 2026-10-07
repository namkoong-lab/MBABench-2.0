"""Read the shared config at <MBABench>/config/config.yaml.

judge/ is a workspace member of MBABench, whose root pyproject.toml exposes
config/python/config.py as the top-level `config` module. That file is the
single home for the local roots (`local.data_root`, `local.output_root`), the
model API keys (`keys.*`) and the LibreOffice binary (`libreoffice_path`), so
nothing secret or machine-specific lives in judge/.

This is a deliberate copy of coding-agents/coding_agent/repo_config.py (kept
separate so each package stays runnable on its own).

RESOLUTION ORDER, first hit wins:

  roots      MBABENCH_DATA_ROOT / MBABENCH_OUTPUT_ROOT, then config local.*,
             then the defaults `data` / `outputs` (relative to the repo root).
  API keys   the environment (OPENAI_API_KEY, ...), then config keys.*.
"""

import logging
import os
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# <repo>/judge/utils/repo_config.py -> <repo>
REPO_ROOT = Path(__file__).resolve().parents[2]

# Points the shared-config lookup at a different directory. An escape hatch
# for layouts where the installed location is wrong. Unset -> Config resolves
# its own directory via the editable install.
REPO_CONFIG_DIR_ENV = "MBABENCH_CONFIG_DIR"

DATA_ROOT_ENV = "MBABENCH_DATA_ROOT"
OUTPUT_ROOT_ENV = "MBABENCH_OUTPUT_ROOT"
DEFAULT_DATA_ROOT = "data"
DEFAULT_OUTPUT_ROOT = "outputs"


def repo_value(*path: str) -> Optional[str]:
    """Non-empty string at a path in <repo>/config/config.yaml, else None.

    * Never raises: a standalone checkout has no config/ directory, so the
      import fails there — and that failure IS the signal to fall through to
      the environment, not an error.
    * Never writes: Config.load() defaults to create_missing=True, which
      seeds a config.yaml from the defaults as a side effect of reading one.
    * A `null` placeholder yields None, so an unset key falls through to the
      next layer rather than resolving to something falsy-but-present.
    """
    try:
        from config import Config
    except ImportError:
        logger.debug("shared `config` module not installed; using env vars")
        return None

    override = os.environ.get(REPO_CONFIG_DIR_ENV)
    # Config.load warns about every unset ${env:VAR} in the file, including
    # keys the judge never reads. Not worth a warning per run.
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
            "could not read the shared config (%s: %s); falling back to "
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


def _root(env_name: str, config_key: str, default: str) -> Path:
    """One local root as an absolute path: env, then config local.*, then default.

    A relative value (from any layer) resolves against the repository root,
    never the working directory, so every judge entry point agrees on where
    `data` and `outputs` are.
    """
    value = (os.environ.get(env_name) or "").strip()
    if not value:
        value = repo_value("local", config_key) or default
    path = Path(value).expanduser()
    return path if path.is_absolute() else (REPO_ROOT / path)


def data_root() -> Path:
    """Where the downloaded benchmark is (`tasks/task_id=<N>/...`)."""
    return _root(DATA_ROOT_ENV, "data_root", DEFAULT_DATA_ROOT)


def output_root() -> Path:
    """Where runs write (`attempts/<label>/...`, `gradings/`)."""
    return _root(OUTPUT_ROOT_ENV, "output_root", DEFAULT_OUTPUT_ROOT)


# Provider -> (environment variable, config/config.yaml keys.* entry).
API_KEYS = {
    "openrouter": ("OPENROUTER_API_KEY", "openrouter_api_key"),
    "gemini": ("GEMINI_API_KEY", "gemini_api_key"),
    "anthropic": ("ANTHROPIC_API_KEY", "anthropic_api_key"),
    "openai": ("OPENAI_API_KEY", "openai_api_key"),
    # TensorBlock Forge gateway (judge provider "tensorblock"): its own key,
    # so a Forge-routed grader can only ever bill Forge credits.
    "forge": ("FORGE_API_KEY", "forge_api_key"),
}


# macOS default install location; `soffice` is not on PATH there by default.
_MACOS_SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def resolve_libreoffice_path() -> Optional[str]:
    """Locate the LibreOffice binary, or None. Same order as the CLI pipeline:

      1. LIBREOFFICE_PATH env var (explicit override)
      2. `libreoffice_path` in <MBABench>/config/config.yaml
      3. `soffice` on PATH (Linux: apt-get install libreoffice-calc)
      4. the macOS app bundle, when it exists

    An explicitly configured path (1 or 2) is returned as-is even if it does
    not exist, so a typo is reported as such instead of silently falling
    through to another binary.
    """
    import shutil
    import sys

    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path:
        return env_path
    repo_path = repo_value("libreoffice_path")
    if repo_path:
        return repo_path
    on_path = shutil.which("soffice")
    if on_path:
        return on_path
    if sys.platform == "darwin" and os.path.exists(_MACOS_SOFFICE):
        return _MACOS_SOFFICE
    return None


def resolve_api_key(provider: str, required: bool = True) -> Optional[str]:
    """API key for `provider`: environment first, then config keys.*.

    Env wins so a session-scoped key never has to be written to disk.
    """
    env_name, cfg_key = API_KEYS[provider]
    key = (os.environ.get(env_name) or "").strip() or repo_value("keys", cfg_key)
    if not key and required:
        raise EnvironmentError(
            f"No {provider} API key: set ${env_name} or keys.{cfg_key} in "
            f"<MBABench>/config/config.yaml"
        )
    return key
