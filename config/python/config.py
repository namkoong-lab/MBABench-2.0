"""config.py — two-tiered config loader for Python.

Reads ``config_default.yaml`` (committed defaults) and ``config.yaml``
(local overrides, gitignored). On first run, ``config.yaml`` is seeded from
the defaults so you have a file to edit. At load time the two are merged:
``config.yaml`` wins, and any key it omits falls back to the defaults.

Usage::

    from config import Config

    cfg = Config.load()                 # uses repo root (../ from this file)
    cfg = Config.load("/some/dir")      # or an explicit config dir

    cfg["app"]["name"]                  # → "my-app"   (dict-style access)
    cfg["app.name"]                     # → "my-app"   (dotted access)
    cfg.get("app.workers")              # → 4          (dotted, with default)
    cfg.get("missing.key", default=0)   # → 0
    cfg.require("database.host")         # → raises KeyError if absent
    cfg.set("app.workers", 8)           # in-memory override (e.g. a CLI flag)
    "app.name" in cfg                   # → True
    cfg.flat_keys()                     # → ["app.name", "app.workers", ...]

Required fields: any value equal to the sentinel ``__REQUIRED__`` marks a
field the user must fill in. Loading aborts with :class:`ConfigRequiredError`
listing every such field, so a half-configured install fails at startup
instead of somewhere deep in a run::

    venv_path: __REQUIRED__     # in config_default.yaml

Pass ``strict=True`` to also treat an unset ``${env:VAR}`` (one with no
``:-default``) as ``__REQUIRED__`` rather than silently expanding to "".

The thin module-level functions ``load_config``, ``get`` and ``deep_merge``
remain as backward-compatible wrappers around :class:`Config`.

Note: this loader expands ``${env:VAR}`` / ``${env:VAR:-default}`` references
(mirroring yaml_parser.sh) but does NOT resolve ``${var}`` cross-references —
those are left as literal strings (the bash loader resolves them).
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Any, Iterator

import yaml

logger = logging.getLogger(__name__)

# Sentinel so ``require`` / ``get`` can distinguish "no default given" from
# an explicit default of ``None``.
_MISSING = object()

# Placeholder for "the user must fill this in". Any field holding this value
# aborts the load. Mirrors CONFIG_REQUIRED in bash/config.sh.
REQUIRED = "__REQUIRED__"


class ConfigRequiredError(ValueError):
    """Raised when fields are still set to the ``__REQUIRED__`` placeholder.

    ``keys`` holds every offending dotted key, so callers can report them all
    rather than fixing one field per run.
    """

    def __init__(self, keys: list[str], user_file: os.PathLike | str | None = None) -> None:
        self.keys = list(keys)
        self.user_file = user_file
        where = f" in {user_file}" if user_file is not None else ""
        listed = "\n".join(f"  - {k}" for k in self.keys)
        super().__init__(
            f"config: {len(self.keys)} field(s) still set to {REQUIRED}{where}:\n{listed}"
        )


class Config:
    """A loaded, merged two-tiered config with dotted access.

    Build instances with :meth:`Config.load`; the constructor takes already
    merged data and is mostly for internal/testing use.

    Attributes
    ----------
    config_dir : Path
        Directory the config was loaded from.
    default_file : Path
        Path to ``config_default.yaml`` (the committed defaults).
    user_file : Path
        Path to ``config.yaml`` (the gitignored local overrides).
    """

    # Repo root is one level up from this file (python/ -> repo root).
    DEFAULT_CONFIG_DIR = Path(__file__).resolve().parent.parent

    DEFAULT_FILENAME = "config_default.yaml"
    USER_FILENAME = "config.yaml"

    # Placeholder marking a field the user must fill in; see module docstring.
    REQUIRED = REQUIRED

    # Matches ${env:VAR} and ${env:VAR:-default}
    _ENV_REF = re.compile(r"\$\{env:([A-Za-z_][A-Za-z0-9_]*)(?::-([^}]*))?\}")

    def __init__(
        self,
        data: dict | None = None,
        *,
        config_dir: str | os.PathLike | None = None,
        default_file: str | os.PathLike | None = None,
        user_file: str | os.PathLike | None = None,
    ) -> None:
        base = Path(config_dir) if config_dir is not None else self.DEFAULT_CONFIG_DIR
        self.config_dir = base
        self.default_file = (
            Path(default_file) if default_file is not None else base / self.DEFAULT_FILENAME
        )
        self.user_file = (
            Path(user_file) if user_file is not None else base / self.USER_FILENAME
        )
        self._data: dict = data if data is not None else {}

    # -- construction ---------------------------------------------------

    @classmethod
    def load(
        cls,
        config_dir: str | os.PathLike | None = None,
        *,
        create_missing: bool = True,
        check_required: bool = True,
        strict: bool = False,
    ) -> "Config":
        """Load the merged two-tiered config.

        Creates ``config.yaml`` from ``config_default.yaml`` on first run
        (unless ``create_missing=False``). Raises :class:`ConfigRequiredError`
        if any field is still ``__REQUIRED__`` (unless ``check_required=False``).
        """
        cfg = cls(config_dir=config_dir)
        cfg.reload(
            create_missing=create_missing,
            check_required=check_required,
            strict=strict,
        )
        return cfg

    def reload(
        self,
        *,
        create_missing: bool = True,
        check_required: bool = True,
        strict: bool = False,
    ) -> "Config":
        """Re-read both tiers from disk and rebuild the merged data in place.

        ``strict`` makes an unset ``${env:VAR}`` with no ``:-default`` expand
        to ``__REQUIRED__`` instead of "", so the required check catches it.
        """
        if not self.default_file.is_file():
            raise FileNotFoundError(f"defaults not found: {self.default_file}")

        # Tier 2: create config.yaml from defaults on first run, dropping
        # whole-line comments (lines whose first non-blank char is #).
        if create_missing and not self.user_file.is_file():
            self._seed_user_file()

        defaults = yaml.safe_load(self.default_file.read_text()) or {}
        user = (
            yaml.safe_load(self.user_file.read_text()) or {}
            if self.user_file.is_file()
            else {}
        )

        merged = self.deep_merge(defaults, user)
        self._data = self._expand_env(merged, strict=strict)
        if check_required:
            self.check_required()
        return self

    def _seed_user_file(self) -> None:
        lines = self.default_file.read_text().splitlines(keepends=True)
        filtered = [ln for ln in lines if not ln.lstrip().startswith("#")]
        while filtered and not filtered[0].strip():
            filtered.pop(0)
        self.user_file.write_text("".join(filtered))
        logger.info("created %s from defaults — edit as needed", self.user_file)

    # -- reads ----------------------------------------------------------

    def get(self, dotted_key: str, default: Any = None) -> Any:
        """Read a nested value by dotted key, e.g. ``cfg.get("app.workers")``."""
        node: Any = self._data
        for part in dotted_key.split("."):
            if not isinstance(node, dict) or part not in node:
                return default
            node = node[part]
        return node

    def require(self, dotted_key: str) -> Any:
        """Like :meth:`get`, but raise ``KeyError`` if the key is absent.

        Use for values your program cannot run without.
        """
        value = self.get(dotted_key, _MISSING)
        if value is _MISSING:
            raise KeyError(f"required config key missing: {dotted_key!r}")
        return value

    def required_keys(self) -> list[str]:
        """Dotted keys still holding the ``__REQUIRED__`` placeholder.

        List elements are reported with their index, e.g. ``features[1]``.
        """
        found: list[str] = []

        def walk(node: Any, prefix: str) -> None:
            if isinstance(node, dict):
                for k, v in node.items():
                    walk(v, f"{prefix}.{k}" if prefix else str(k))
            elif isinstance(node, list):
                for i, v in enumerate(node):
                    walk(v, f"{prefix}[{i}]")
            elif node == REQUIRED:
                found.append(prefix)

        walk(self._data, "")
        return found

    def check_required(self) -> "Config":
        """Raise :class:`ConfigRequiredError` if any field is ``__REQUIRED__``."""
        pending = self.required_keys()
        if pending:
            raise ConfigRequiredError(pending, self.user_file)
        return self

    def flat_keys(self) -> list[str]:
        """All leaf keys as dotted paths (mirrors bash ``config keys``)."""
        keys: list[str] = []

        def walk(node: Any, prefix: str) -> None:
            if isinstance(node, dict):
                for k, v in node.items():
                    walk(v, f"{prefix}.{k}" if prefix else str(k))
            else:
                keys.append(prefix)

        walk(self._data, "")
        return keys

    def as_dict(self, *, copy: bool = True) -> dict:
        """Return the merged config as a plain dict (deep-copied by default)."""
        import copy as _copy

        return _copy.deepcopy(self._data) if copy else self._data

    def to_yaml(self) -> str:
        """Serialize the merged config back to YAML."""
        return yaml.safe_dump(self._data, sort_keys=False)

    # -- writes (in-memory only; never touches disk) --------------------

    def set(self, dotted_key: str, value: Any) -> "Config":
        """Override a value in memory by dotted key, creating parents as needed.

        Useful for layering CLI flags / runtime values on top of the files.
        Does NOT write to ``config.yaml``.
        """
        parts = dotted_key.split(".")
        node = self._data
        for part in parts[:-1]:
            child = node.get(part)
            if not isinstance(child, dict):
                child = {}
                node[part] = child
            node = child
        node[parts[-1]] = value
        return self

    # -- dunders: behave like a read-only mapping -----------------------

    def __getitem__(self, dotted_key: str) -> Any:
        value = self.get(dotted_key, _MISSING)
        if value is _MISSING:
            raise KeyError(dotted_key)
        return value

    def __contains__(self, dotted_key: object) -> bool:
        return isinstance(dotted_key, str) and self.get(dotted_key, _MISSING) is not _MISSING

    def __iter__(self) -> Iterator[str]:
        return iter(self._data)

    def __len__(self) -> int:
        return len(self._data)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Config):
            return self._data == other._data
        if isinstance(other, dict):
            return self._data == other
        return NotImplemented

    def __repr__(self) -> str:
        return f"Config(config_dir={str(self.config_dir)!r}, keys={list(self._data)})"

    # -- static helpers -------------------------------------------------

    @staticmethod
    def deep_merge(base: dict, override: dict) -> dict:
        """Recursively merge ``override`` onto ``base``; ``override`` wins.

        Nested dicts are merged key-by-key. Any non-dict value (including
        lists) is replaced wholesale by the override.
        """
        result = dict(base)
        for key, value in override.items():
            if (
                key in result
                and isinstance(result[key], dict)
                and isinstance(value, dict)
            ):
                result[key] = Config.deep_merge(result[key], value)
            else:
                result[key] = value
        return result

    @classmethod
    def _expand_env(cls, value: Any, *, strict: bool = False) -> Any:
        """Recursively expand ${env:VAR} / ${env:VAR:-default} in string leaves."""
        if isinstance(value, dict):
            return {k: cls._expand_env(v, strict=strict) for k, v in value.items()}
        if isinstance(value, list):
            return [cls._expand_env(v, strict=strict) for v in value]
        if isinstance(value, str):
            return cls._expand_env_str(value, strict=strict)
        return value

    @classmethod
    def _expand_env_str(cls, text: str, *, strict: bool = False) -> str:
        def repl(match: re.Match) -> str:
            name, has_default = match.group(1), match.group(2)
            if name in os.environ:
                return os.environ[name]
            if has_default is not None:
                return has_default
            logger.debug("env var not set: %s", name)
            # Under strict, hand the required check a value it will reject
            # instead of quietly substituting an empty string.
            return REQUIRED if strict else ""

        return cls._ENV_REF.sub(repl, text)


# -- backward-compatible module-level API ------------------------------
# Thin wrappers so existing callers keep working unchanged.


def load_config(
    config_dir: str | os.PathLike | None = None,
    *,
    check_required: bool = True,
    strict: bool = False,
) -> Config:
    """Load and return a :class:`Config` (replaces the old dict-returning API).

    A ``Config`` supports ``cfg["app"]["name"]`` and ``get(cfg, "app.name")``,
    so existing callers keep working.
    """
    return Config.load(config_dir, check_required=check_required, strict=strict)


def get(cfg: Config | dict, dotted_key: str, default: Any = None) -> Any:
    """Read a nested value by dotted key from a ``Config`` or plain dict."""
    if isinstance(cfg, Config):
        return cfg.get(dotted_key, default)
    node: Any = cfg
    for part in dotted_key.split("."):
        if not isinstance(node, dict) or part not in node:
            return default
        node = node[part]
    return node


def deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge ``override`` onto ``base``; ``override`` wins."""
    return Config.deep_merge(base, override)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    import json

    print(json.dumps(load_config().as_dict(), indent=2))
