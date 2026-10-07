"""Run configuration: one YAML file fully describes a run.

The config names its cohort with `agent_model_name`; the entry in
agent_identities.yaml supplies cli/model/effort/extra_args/env (see
agent_identity.py). `template_version` picks the prompt (default v13), `tasks`
bounds the sweep range, `sandbox` / `limits` are run settings.

Tasks come from <data_root>/tasks/ and attempts go to
<output_root>/attempts/<agent_model_name>/ — repo_config.data_root() /
output_root(), i.e. local.* in <MBABench>/config/config.yaml or the
MBABENCH_DATA_ROOT / MBABENCH_OUTPUT_ROOT environment variables.

Secrets are never stored in run configs: the agent's API key comes from the
environment (or a local .env next to this package), falling back to
config/config.yaml keys.*.
"""
import hashlib
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse

import yaml

from . import repo_config
from .agent_identity import AgentIdentity, resolve_agent_identity

PACKAGE_DIR = Path(__file__).resolve().parent
PROMPTS_DIR = PACKAGE_DIR / "prompts"

# The trajectory relay runs from the repo's copy, bind-mounted read-only over
# the one baked into the image (sandbox.run_in_sandbox), so a relay fix never
# needs a new image tag: the tag is the recorded CLI pin. Rows record the
# mounted file's hash (extra_configs.relay).
RELAY_SOURCE = PACKAGE_DIR.parent / "docker" / "traj_relay.py"
RELAY_TARGET = "/usr/local/bin/traj_relay.py"

# Codex model catalog for model ids Codex does not know (the TensorBlock Forge
# cohorts). Codex gives an unknown id "fallback metadata" with a 272,000-token
# window it will not raise, so it compacts at ~245k however large the model's
# real window is. Each entry here reproduces Codex's fallback metadata exactly
# and changes only context_window / max_context_window to the model's own
# limit. Mounted read-only when an identity names it in extra_args
# (-c model_catalog_json=...); rows record the file's hash
# (extra_configs.codex_model_catalog).
CODEX_CATALOG_SOURCE = PACKAGE_DIR.parent / "docker" / "codex_model_catalog.json"
CODEX_CATALOG_TARGET = "/etc/codex/model_catalog.json"

# Env var name, and the config/config.yaml keys.* fallback, per agent CLI.
AGENT_KEY_ENV = {"claude": "ANTHROPIC_API_KEY", "codex": "OPENAI_API_KEY"}
AGENT_KEY_CONFIG = {"claude": "anthropic_api_key", "codex": "openai_api_key"}

# A TensorBlock Forge identity (its env.TRAJ_UPSTREAM points the relay at the
# gateway) is keyed on its own and never falls back to the vendor key — that
# fallback would send the OpenAI key to TensorBlock. The key still enters the
# container under the CLI's usual name (OPENAI_API_KEY): `codex login` and
# the traj provider's env_key read it.
FORGE_KEY_ENV = "FORGE_API_KEY"
FORGE_KEY_CONFIG = "forge_api_key"

# Egress allowlist per agent CLI: the model API only — CLI telemetry is
# disabled via env, and the firewall fails closed on unresolvable domains.
# Extend per-run via sandbox.network_allow if a CLI needs another endpoint.
DEFAULT_ALLOWED_DOMAINS = {
    "claude": ["api.anthropic.com"],
    "codex": ["api.openai.com"],
}

# Keys older run configs carried that now live elsewhere. Refused (not
# ignored) so a stale config gets migrated deliberately.
STALE_KEYS = {
    "identity": "renamed to agent_model_name (a label registered in agent_identities.yaml)",
    "agent": "pinned by the agent_model_name entry in agent_identities.yaml",
    "internal": "there is no object store; attempts go to <output_root>/attempts/",
    "workspaces_dir": "attempt working dirs live under coding-agents/workspaces/",
}
# Keys from earlier layouts that are accepted when they carry the one value
# this runner implements, and refused with a message otherwise.
LEGACY_KEYS = {
    "mode": "internal",   # a benchmark task; external task folders are gone
    "benchmark": "v2",    # the benchmark is MBABench, always
    "source": "local",    # tasks come from <data_root>
    "sink": "local",      # attempts go to <output_root>
}


def load_dotenv_if_present(path: Path | None = None) -> None:
    """Tiny .env loader (KEY=VALUE lines); never overrides existing env."""
    env_path = path or (PACKAGE_DIR.parent / ".env")
    if not env_path.exists():
        return
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip("'\"")
        if key and key not in os.environ:
            os.environ[key] = value


@dataclass
class AgentConfig:
    cli: str  # "claude" | "codex"
    model: str
    effort: str | None = None
    extra_args: list[str] = field(default_factory=list)
    env: dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_identity(cls, identity: AgentIdentity) -> "AgentConfig":
        return cls(cli=identity.cli, model=identity.model, effort=identity.effort,
                   extra_args=list(identity.extra_args), env=dict(identity.env))


@dataclass
class SandboxConfig:
    mode: str = "docker"  # "docker" | "host" (host = UNSANDBOXED, dev only)
    # The tag is the CLI-version pin recorded per attempt
    # (extra_configs.sandbox_image); retag any rebuild that changes contents.
    image: str = "mbabench-coding-agent:v2"
    network_allow: list[str] = field(default_factory=list)
    cpus: int = 4
    memory: str = "8g"


@dataclass
class LimitsConfig:
    wall_clock_seconds: int = 14400  # 4h
    # A success faster than this many seconds is held as needs_review instead
    # of recorded (a guard for junk deliveries by slow frontier cohorts). Off by
    # default: small or low-effort models legitimately finish a task in under a
    # minute. Re-enable per run with `limits: {junk_seconds: 180}`.
    junk_seconds: int = 0
    # When true, time the relay spends waiting out provider 429 / quota
    # refusals (its upstream_retries delays) does not count against the wall
    # clock - the agent still gets wall_clock_seconds of working time. Off by
    # default; rows record it in extra_configs.limits when on.
    exclude_provider_waits: bool = False


@dataclass
class TaskRange:
    """The inclusive task-id window a sweep walks (`tasks:` in a run config).
    The sweep applies "present under <data_root>/tasks/, ascending id" on top;
    this only bounds the ids."""
    first: int = 1
    last: int = 101

    def __post_init__(self):
        if self.first < 1 or self.last < self.first:
            raise ValueError(f"tasks must be an ascending 1-based range, got {self.first}-{self.last}")

    def __contains__(self, task_id: int) -> bool:
        return self.first <= task_id <= self.last

    def __iter__(self):
        return iter(range(self.first, self.last + 1))

    def __str__(self) -> str:
        return f"{self.first}-{self.last}"


DEFAULT_TEMPLATE_VERSION = "v13"
TEMPLATE_VERSIONS = ("v8", "v9", "v10", "v11", "v12", "v13", "v14", "v15")

# Files a template promises the agent, repo-root-relative. They are seeded
# into starting_files/ beside the task inputs and snapshotted with the prompt
# files. Declared on the template — never in a run config — so the recorded
# prompt_version and the file the agent saw cannot disagree. v11/v13 deliver
# the standards differently — staged into the WORKSPACE ROOT as
# HOUSE_STANDARDS.md from prompts/house_standards_v1.md (a byte-identical
# copy of <MBABench>/house_standards/House_Standards_v1.md); see
# prompt_builder.TEMPLATE_EXTRAS. Both routes record house_standards below.
TEMPLATE_ATTACHMENTS = {
    "v12": ["house_standards/House_Standards_v1.md"],
}
_HOUSE_STANDARDS_RE = re.compile(r"^House_Standards_v(\d+)\.md$")


@dataclass
class RunConfig:
    agent_model_name: str  # cohort label (attempt rows' agent_model_name, the attempts/ folder)
    identity: AgentIdentity
    agent: AgentConfig
    record_trajectory: bool = True  # per-step API request/response capture (docker mode only)
    system_prompt: str = "system_prompt_coding_v1.txt"
    template_version: str = DEFAULT_TEMPLATE_VERSION  # v13 = rubric-free + house standards (default); v12 = v9 + house standards (rubric-bearing); v14 = v9 text (rubric, no standards); v15 = v10 text (neither); v10/v11 = rubric-scrubbed experiment; v9 = Questions-sheet mirror; v8 = rubric mirror
    sandbox: SandboxConfig = field(default_factory=SandboxConfig)
    limits: LimitsConfig = field(default_factory=LimitsConfig)
    tasks: "TaskRange | None" = None  # the sweep range (run_sweep)
    workspaces_dir: Path = PACKAGE_DIR.parent / "workspaces"
    config_path: Path | None = None  # the YAML this was loaded from (copied into the attempt dir)

    @property
    def api_key_env(self) -> str:
        return AGENT_KEY_ENV[self.agent.cli]

    @property
    def allowed_domains(self) -> list[str]:
        # A Forge run reaches TensorBlock only: without the vendor's API in the
        # allowlist, nothing in the container can send the Forge key there.
        if is_forge(self.agent):
            host = urlparse(self.agent.env["TRAJ_UPSTREAM"]).hostname
            base = [host] if host else []
        else:
            base = DEFAULT_ALLOWED_DOMAINS[self.agent.cli]
        return list(dict.fromkeys(base + self.sandbox.network_allow))

    @property
    def data_root(self) -> Path:
        """Where the benchmark tasks are read from."""
        return repo_config.data_root().resolve()

    @property
    def output_root(self) -> Path:
        """Where attempts are written (attempts/<agent_model_name>/...)."""
        return repo_config.output_root().resolve()

    def extra_configs(self) -> dict:
        """What an attempt row's extra_configs records: the identity's pinned
        settings, the sandbox image (it pins the CLI version), the harness
        defaults applied under the identity (agents.harness_defaults) and,
        when the template ships house standards, which text the agent saw."""
        out = {**self.identity.extra_configs(), "sandbox_image": self.sandbox.image}
        out.update(house_standards_provenance(self))
        from .agents import harness_defaults  # local: agents imports this module
        relay = self.record_trajectory and self.sandbox.mode == "docker"
        out["harness_defaults"] = harness_defaults(self.agent, relay)
        if relay and RELAY_SOURCE.is_file():
            out["relay"] = {"source": "docker/traj_relay.py",
                            "sha256": hashlib.sha256(RELAY_SOURCE.read_bytes()).hexdigest()}
        if self.limits.exclude_provider_waits:
            out["limits"] = {"wall_clock_seconds": self.limits.wall_clock_seconds,
                             "exclude_provider_waits": True}
        if uses_codex_catalog(self.agent) and CODEX_CATALOG_SOURCE.is_file():
            out["codex_model_catalog"] = {
                "source": "docker/codex_model_catalog.json",
                "sha256": hashlib.sha256(CODEX_CATALOG_SOURCE.read_bytes()).hexdigest()}
        return out


def template_attachments(cfg: RunConfig) -> list[Path]:
    """Absolute paths of the template's declared attachments (empty for
    templates that declare none).

    Raises FileNotFoundError when the repo root or a file is missing: a
    template whose directive names a file the agent will never find must
    not run, and the caller treats this as infra_failure (no row).
    """
    rels = TEMPLATE_ATTACHMENTS.get(cfg.template_version, [])
    if not rels:
        return []
    root = repo_config.monorepo_root()
    if root is None:
        raise FileNotFoundError(
            f"template {cfg.template_version} needs {rels} from the MBABench "
            f"root, which could not be located (install the workspace, or "
            f"check out coding-agents inside the repo)"
        )
    paths = [root / rel for rel in rels]
    missing = [str(p) for p in paths if not p.is_file()]
    if missing:
        raise FileNotFoundError(
            f"template {cfg.template_version} declares attachment(s) that do "
            f"not exist: {missing}"
        )
    return paths


def house_standards_provenance(cfg: RunConfig) -> dict:
    """{"house_standards": {version, file, sha256[, delivered_as]}} for the
    standards file the template delivers, else {}. The hash is taken at run
    time so the row records the text actually delivered, not the version the
    name claims. Two delivery routes: a declared attachment seeded into
    starting_files/ (v12), or a workspace-root extra (v11/v13 stage
    prompts/house_standards_v1.md as HOUSE_STANDARDS.md — `delivered_as`
    records that name; `file` keeps the versioned source name)."""
    for path in template_attachments(cfg):
        m = _HOUSE_STANDARDS_RE.match(path.name)
        if m:
            return {"house_standards": {
                "version": int(m.group(1)),
                "file": path.name,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            }}
    from .prompt_builder import prompt_extra_paths  # local: prompt_builder imports this module
    for src, ws_name in prompt_extra_paths(cfg):
        m = _HOUSE_STANDARDS_RE.match(src.name.replace("house_standards_v", "House_Standards_v", 1))
        if m:
            return {"house_standards": {
                "version": int(m.group(1)),
                "file": f"House_Standards_v{m.group(1)}.md",
                "delivered_as": ws_name,
                "sha256": hashlib.sha256(src.read_bytes()).hexdigest(),
            }}
    return {}


def load_config(path: str | Path) -> RunConfig:
    path = Path(path)
    if not path.is_file():
        hint = ""
        candidate = Path(__file__).resolve().parents[1] / "run_configs" / path.name
        if candidate.is_file():
            hint = f" (did you mean run_configs/{path.name}?)"
        raise SystemExit(f"run config not found: {path}{hint}")
    raw = yaml.safe_load(path.read_text()) or {}
    stale = [k for k in STALE_KEYS if k in raw]
    if stale:
        raise ValueError(
            "run config carries key(s) that no longer belong there:\n"
            + "\n".join(f"  - {k}: {STALE_KEYS[k]}" for k in stale)
        )
    for key, only in LEGACY_KEYS.items():
        if key in raw and str(raw[key]).lower() != only:
            raise ValueError(
                f"run config key `{key}` can only be '{only}' (it may also be omitted): "
                f"the runner reads tasks from <data_root> and writes attempts to "
                f"<output_root>, nothing else"
            )

    identity = resolve_agent_identity(raw)
    cfg = RunConfig(
        agent_model_name=identity.agent_model_name,
        identity=identity,
        agent=AgentConfig.from_identity(identity),
        record_trajectory=bool(raw.get("record_trajectory", True)),
        system_prompt=raw.get("system_prompt", "system_prompt_coding_v1.txt"),
        template_version=raw.get("template_version", DEFAULT_TEMPLATE_VERSION),
        sandbox=SandboxConfig(**(raw.get("sandbox") or {})),
        limits=LimitsConfig(**(raw.get("limits") or {})),
        tasks=TaskRange(**raw["tasks"]) if raw.get("tasks") else None,
        config_path=path.resolve(),
    )
    if cfg.template_version not in TEMPLATE_VERSIONS:
        raise ValueError(f"template_version must be one of {', '.join(TEMPLATE_VERSIONS)}")
    if cfg.sandbox.mode not in ("docker", "host"):
        raise ValueError('sandbox.mode must be "docker" or "host"')
    return cfg


def is_forge(agent: AgentConfig) -> bool:
    """True when the identity sends its calls to TensorBlock Forge."""
    return "tensorblock" in agent.env.get("TRAJ_UPSTREAM", "").lower()


def uses_codex_catalog(agent: AgentConfig) -> bool:
    """True when the identity points Codex at the mounted model catalog."""
    return agent.cli == "codex" and any(CODEX_CATALOG_TARGET in a for a in agent.extra_args)


def resolve_api_key(cfg: RunConfig) -> str:
    """The agent's API key: environment first, then config/config.yaml keys.*.
    A Forge identity resolves the Forge key or nothing (see FORGE_KEY_ENV)."""
    if is_forge(cfg.agent):
        return (os.environ.get(FORGE_KEY_ENV)
                or repo_config.repo_value("keys", FORGE_KEY_CONFIG)
                or "")
    return (os.environ.get(cfg.api_key_env)
            or repo_config.repo_value("keys", AGENT_KEY_CONFIG[cfg.agent.cli])
            or "")


def resolve_secrets(cfg: RunConfig) -> str:
    """Fail fast on a missing API key, before any work is done."""
    forge = is_forge(cfg.agent)
    # Only the relay reads TRAJ_UPSTREAM. Without it the CLI would call its
    # vendor's API directly — carrying the Forge key.
    if forge and not (cfg.record_trajectory and cfg.sandbox.mode == "docker"):
        raise SystemExit(
            f"{cfg.agent_model_name} reaches TensorBlock Forge only through the "
            f"trajectory relay: it needs sandbox.mode docker and record_trajectory on"
        )
    api_key = resolve_api_key(cfg)
    if not api_key:
        if forge:
            raise SystemExit(
                f"Missing {FORGE_KEY_ENV}: set it in the environment, a .env next "
                f"to coding_agent/, or <MBABench>/config/config.yaml "
                f"keys.{FORGE_KEY_CONFIG} (a Forge identity never falls back to "
                f"{cfg.api_key_env})"
            )
        raise SystemExit(
            f"Missing {cfg.api_key_env}: set it in the environment, a .env next "
            f"to coding_agent/, or <MBABench>/config/config.yaml "
            f"keys.{AGENT_KEY_CONFIG[cfg.agent.cli]}"
        )
    # The vendors' key prefixes differ (Anthropic: sk-ant-..., OpenAI: sk-... but
    # never sk-ant-). A key in the wrong slot would otherwise be retried 15 times
    # against the wrong API as 401 Unauthorized before the attempt fails.
    if not forge:
        looks_anthropic = api_key.startswith("sk-ant-")
        if cfg.agent.cli == "codex" and looks_anthropic:
            raise SystemExit(
                f"{cfg.api_key_env} holds an Anthropic key (sk-ant-...); Codex needs an "
                f"OpenAI key — check keys.openai_api_key in <MBABench>/config/config.yaml"
            )
        if cfg.agent.cli == "claude" and api_key.startswith("sk-") and not looks_anthropic:
            raise SystemExit(
                f"{cfg.api_key_env} does not look like an Anthropic key (sk-ant-...); "
                f"check keys.anthropic_api_key in <MBABench>/config/config.yaml"
            )
    return api_key
