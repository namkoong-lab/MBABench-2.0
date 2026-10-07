import argparse
import os
from pathlib import Path

import yaml

try:
    from . import repo_config
except ImportError:  # imported as a bare module (utils/ on sys.path)
    import repo_config

# judge/ root; project_configs.yaml (tracked, no secrets) lives here.
JUDGE_ROOT = Path(__file__).resolve().parent.parent
_CONFIG_PATH = JUDGE_ROOT / "project_configs.yaml"

# The benchmark this judge grades. There is one: the 101-task MBABench
# (historically "v2"); the name is kept where rows and logs record it.
BENCHMARK = "v2"

# The rubric pair and category order the benchmark is graded with. Exported
# by load_project_configs() as {PREFIX}_JUDGE_RUBRIC, _JUDGE_RUBRIC_VERSION,
# _JUDGE_RUBRIC_WEIGHT, _JUDGE_RUBRIC_WEIGHT_VERSION and _JUDGE_CHECK_ORDER
# (paths relative to judge/).
RUBRIC = "prompts/rubrics/rubric_9.json"
RUBRIC_VERSION = "9"
RUBRIC_WEIGHT = "prompts/rubrics/rubric_9_weights.json"
RUBRIC_WEIGHT_VERSION = "9"
CHECK_ORDER = (
    "Accuracy,Assumptions,Documentation,Error Checks,Flexibility,"
    "Formatting,Formulas,Model Outputs & Executive Summary,"
    "Potential Dangers,Purpose & Scope,Rounding,Structure"
)


def get_absolute_path(path) -> str:
    """Convert a path to an absolute path."""
    return str(Path(path).resolve())


def relative_path_from_project_root(path) -> str:
    """Convert a path to be relative from the project root directory."""
    project_root = Path(__file__).parent.parent.resolve()
    path = Path(path)

    # interpret .. or . as relative to project root and resolve to absolute path
    absolute_path = (project_root / path).resolve()
    return absolute_path


def _flatten_dict(d, prefix=""):
    """Recursively flatten a nested dict into {PREFIX_KEY: value} pairs, skipping None values."""
    items = {}
    for key, value in d.items():
        full_key = f"{prefix}_{key.upper()}" if prefix else key.upper()
        if isinstance(value, dict):
            items.update(_flatten_dict(value, full_key))
        elif value is not None:
            items[full_key] = value
    return items


def _read_config():
    with open(_CONFIG_PATH) as f:
        return yaml.safe_load(f) or {}


_PREFIX = None


def project_prefix() -> str:
    """Env-var prefix, project.name upper-cased (e.g. MBABENCHJUDGE)."""
    global _PREFIX
    if _PREFIX is None:
        _PREFIX = _read_config().get("project", {}).get("name", "").upper()
    return _PREFIX


def _rubric_env(prefix):
    """The {PREFIX}_* vars pinning the rubric pair and category order."""
    return {
        f"{prefix}_BENCHMARK": BENCHMARK,
        f"{prefix}_JUDGE_RUBRIC": RUBRIC,
        f"{prefix}_JUDGE_RUBRIC_VERSION": RUBRIC_VERSION,
        f"{prefix}_JUDGE_RUBRIC_WEIGHT": RUBRIC_WEIGHT,
        f"{prefix}_JUDGE_RUBRIC_WEIGHT_VERSION": RUBRIC_WEIGHT_VERSION,
        f"{prefix}_JUDGE_CHECK_ORDER": CHECK_ORDER,
    }


def load_project_configs(verbose=False):
    """Load project_configs.yaml (+ the rubric constants above) into environment variables.

    Env var names follow: {PROJECT_NAME}_{SECTION}_{...}_{KEY}
    where PROJECT_NAME comes from project.name in the config.
    """
    if verbose:
        print("*" * 126)
        print(f"Loading project configs from {_CONFIG_PATH}...")
    config = _read_config()
    prefix = project_prefix()

    loaded_configs = {}
    for section_name, section in config.items():
        if not isinstance(section, dict):
            continue
        section_prefix = f"{prefix}_{section_name.upper()}"
        for env_key, value in _flatten_dict(section, section_prefix).items():
            if section_name == "paths":
                # Relative paths are relative to judge/, whatever the cwd.
                value = str((JUDGE_ROOT / str(value)).resolve())
            loaded_configs[env_key] = value

    # paths.libreoffice_path left null in project_configs.yaml means "find it":
    # LIBREOFFICE_PATH, config/config.yaml libreoffice_path, soffice on PATH,
    # the macOS app bundle (repo_config.resolve_libreoffice_path). An explicit
    # value in the yaml is kept as written.
    lo_key = f"{prefix}_PATHS_LIBREOFFICE_PATH"
    if lo_key not in loaded_configs:
        found = repo_config.resolve_libreoffice_path()
        if found:
            loaded_configs[lo_key] = found

    loaded_configs.update(_rubric_env(prefix))

    for env_key, value in loaded_configs.items():
        if verbose:
            print(f"Setting env var {env_key} = {value}")
        os.environ[env_key] = str(value)
    if verbose:
        print("*" * 126)
    return loaded_configs, prefix


LIBREOFFICE_HELP = (
    "No LibreOffice binary found. Install LibreOffice (Linux: apt-get install "
    "libreoffice-calc; macOS: LibreOffice.app) and, if it is not `soffice` on "
    "PATH, point `libreoffice_path` in <MBABench>/config/config.yaml, the "
    "LIBREOFFICE_PATH environment variable, or `paths.libreoffice_path` in "
    "judge/project_configs.yaml at the soffice binary."
)


def libreoffice_path(required=True) -> str:
    """The soffice binary the judge runs (resolved by load_project_configs), or ""
    when none was found and `required` is False. Raises EnvironmentError with
    install guidance otherwise."""
    value = load_env_var("PATHS_LIBREOFFICE_PATH", default="") or ""
    if not value and required:
        raise EnvironmentError(LIBREOFFICE_HELP)
    return str(value)


def load_env_var(var_name: str, default=None, prefix=None, required=False):
    """Helper to load an env var with optional default."""
    if prefix is None:
        prefix = project_prefix()
    var_name = f"{prefix}_{var_name.upper()}"
    value = os.environ.get(var_name, None)

    if value is None:
        if not required:
            from .logger import logger

            logger.debug(
                f"Environment variable {var_name} not set. Using default: {default}"
            )
            value = default
        else:
            raise EnvironmentError(f"Required environment variable {var_name} not set.")
    return value


### YAML dump helpers for conversation messages
class _LiteralStr(str):
    pass


def _literal_str_representer(dumper, data):
    return dumper.represent_scalar("tag:yaml.org,2002:str", data, style="|")


yaml.add_representer(_LiteralStr, _literal_str_representer)
yaml.add_representer(_LiteralStr, _literal_str_representer, Dumper=yaml.SafeDumper)


def _blockify_multiline_strings(obj):
    if isinstance(obj, str):
        if "\n" in obj:
            # PyYAML silently falls back to quoted style if any line has trailing whitespace.
            return _LiteralStr("\n".join(line.rstrip() for line in obj.split("\n")))
        return obj
    if isinstance(obj, list):
        return [_blockify_multiline_strings(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _blockify_multiline_strings(v) for k, v in obj.items()}
    return obj


def dump_messages_yaml(messages, path):
    """Write conversation messages as YAML, using literal block scalars for multi-line strings."""
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            _blockify_multiline_strings(messages),
            f,
            sort_keys=False,
            allow_unicode=True,
            width=10_000,
            default_flow_style=False,
        )


### Argparser helper
def str2bool(v):
    """Convert string to boolean for argparse."""
    if isinstance(v, bool):
        return v
    if v.lower() in ("yes", "true", "t", "y", "1"):
        return True
    elif v.lower() in ("no", "false", "f", "n", "0"):
        return False
    else:
        raise argparse.ArgumentTypeError("Boolean value expected.")


if __name__ == "__main__":
    load_project_configs()
