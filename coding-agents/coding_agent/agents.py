"""Per-CLI invocation: how to launch each coding agent headless.

Both agents receive PROMPT.md on stdin and must emit machine-readable
transcripts on stdout (captured by the sandbox layer). Exact flags for model /
effort pinning are verified empirically at rung 0; agent.extra_args in the run
config exists so rung-0 corrections never require code changes.
"""
from .config import AgentConfig

# ---------------------------------------------------------------------------
# Harness defaults for long thinking turns (2026-09-11). Applied to EVERY run
# of each CLI, underneath the identity's own env/args (an identity entry can
# still override a key). Recorded per row as extra_configs.harness_defaults
# so an attempt says what it ran with even if these change later.
#
# Why: frontier models at max effort can think silently for >5 min on one
# step. Both CLIs ship a ~5-minute stream idle watchdog that then aborts the
# request as "broken", retries the same turn (which stalls the same way) and
# finally gives up — the 2026-09-10 Fable 5.1 run lost 4 of 10 tasks to this
# with no workbook written. Claude Code also falls back to NON-streaming
# requests capped at max_tokens 64000, which can never finish a long turn.
#
# Claude Code (code.claude.com/docs/en/network-config.md, errors.md):
#   CLAUDE_STREAM_IDLE_TIMEOUT_MS       event-level watchdog, default 300 s
#   CLAUDE_BYTE_STREAM_IDLE_TIMEOUT_MS  byte-level watchdog, default 180 s
#                                       (clamped to 30 min — the max we set)
#   CLAUDE_CODE_DISABLE_NONSTREAMING_FALLBACK=1  keep retrying as streams
#   CLAUDE_CODE_MAX_RETRIES             default 10, cap 15
# Codex (config reference, model_providers.<id>):
#   stream_idle_timeout_ms  default 300000;  stream_max_retries default 5;
#   request_max_retries default 4 — passed as -c overrides on whichever
#   provider the run uses (the traj relay provider, or the built-in openai).
# 2026-09-11 retry of the four stalled tasks (rows 1456/1457/1459/1460): with
# the four values above every run died at ~5.5 min into a long turn with
# "API Error: The operation timed out." and ZERO retries. That message is the
# Bun runtime's hardcoded 5-minute fetch timeout (oven-sh/bun#16682), not a
# Claude Code watchdog. Claude Code passes `timeout: false` to fetch only when
# its byte watchdog is armed, and the byte watchdog is armed only when
# ANTHROPIC_BASE_URL is api.anthropic.com — under the traj relay
# (http://127.0.0.1:9877) it is not, so the runtime timeout stays live.
# API_FORCE_IDLE_TIMEOUT=0 makes Claude Code pass `timeout: false` regardless
# (read from the 2.1.251 binary: fetchOptions builder `Ci`, gate
# `!truthy(t) && (hasBodyIdleWatchdog || falsy(t))`). The 30-min event
# watchdog and the SDK's API_TIMEOUT_MS still bound a truly dead stream.
CLAUDE_HARNESS_ENV = {
    "CLAUDE_STREAM_IDLE_TIMEOUT_MS": "1800000",
    "CLAUDE_BYTE_STREAM_IDLE_TIMEOUT_MS": "1800000",
    "CLAUDE_CODE_DISABLE_NONSTREAMING_FALLBACK": "1",
    "CLAUDE_CODE_MAX_RETRIES": "15",
    "API_FORCE_IDLE_TIMEOUT": "0",
}
CODEX_HARNESS_PROVIDER_CONFIG = {
    "stream_idle_timeout_ms": "1800000",
    "stream_max_retries": "15",
    "request_max_retries": "15",
}


def codex_provider_id(relay: bool) -> str:
    return "traj" if relay else "openai"


def harness_defaults(agent: AgentConfig, relay: bool = False) -> dict:
    """What the harness adds underneath the identity, for extra_configs."""
    if agent.cli == "claude":
        return {"env": dict(CLAUDE_HARNESS_ENV)}
    if agent.cli == "codex":
        pid = codex_provider_id(relay)
        return {"codex_config": {f"model_providers.{pid}.{k}": v
                                 for k, v in CODEX_HARNESS_PROVIDER_CONFIG.items()}}
    return {}


def build_command(agent: AgentConfig, relay: bool = False) -> list[str]:
    if agent.cli == "claude":
        cmd = [
            "claude",
            "-p",  # print (headless) mode; prompt read from stdin
            "--output-format", "stream-json",
            "--verbose",
            "--model", agent.model,
            "--dangerously-skip-permissions",  # the container is the safety boundary
        ]
        if agent.effort:
            cmd += ["--effort", agent.effort]  # verified in Claude Code 2.1.223 (low|medium|high|xhigh|max)
    elif agent.cli == "codex":
        import shlex
        args = [
            "codex", "exec",
            "--json",
            "--model", agent.model,
            "--skip-git-repo-check",
            "--dangerously-bypass-approvals-and-sandbox",  # container is the boundary
        ]
        if agent.effort:
            args += ["-c", f"model_reasoning_effort={agent.effort}"]
        if relay:
            # codex ignores OPENAI_BASE_URL; route through the trajectory relay
            # via a custom provider (env_key keeps API-key billing).
            args += ["-c", "model_providers.traj.name=traj-relay",
                     "-c", 'model_providers.traj.base_url="http://127.0.0.1:9877/v1"',
                     "-c", "model_providers.traj.wire_api=responses",
                     "-c", "model_providers.traj.env_key=OPENAI_API_KEY",
                     "-c", "model_provider=traj"]
        for key, value in harness_defaults(agent, relay)["codex_config"].items():
            args += ["-c", f"{key}={value}"]  # identity extra_args come after, so they win
        args += list(agent.extra_args) + ["-"]  # read prompt from stdin
        # codex 0.146 does not read OPENAI_API_KEY implicitly: an explicit
        # `codex login --with-api-key` (stdin) must store it first. The login
        # pipeline consumes only its own pipe; the exec'd codex inherits the
        # outer stdin (PROMPT.md). auth.json lives in the container's
        # ephemeral HOME and dies with it.
        cmd = ["bash", "-c",
               "printenv OPENAI_API_KEY | codex login --with-api-key >/dev/null 2>&1 && exec "
               + shlex.join(args)]
        return cmd
    else:
        raise ValueError(f"Unknown agent cli: {agent.cli}")
    return cmd + list(agent.extra_args)


def agent_env(agent: AgentConfig, api_key_env: str, api_key: str) -> dict:
    env = {
        api_key_env: api_key,
        # Reduce nonessential phoning-home where the CLIs support it.
        "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC": "1",
        "DISABLE_TELEMETRY": "1",
        "DISABLE_ERROR_REPORTING": "1",
    }
    if agent.cli == "claude":
        env.update(CLAUDE_HARNESS_ENV)  # harness defaults for long thinking turns
    env.update(agent.env)  # the identity's own pins win
    return env
