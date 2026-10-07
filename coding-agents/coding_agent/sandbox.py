"""Sandbox: run the agent command in an isolated container (or, for rung-0
debugging only, directly on the host — UNSANDBOXED).

Docker mode, per attempt:
  - fresh container from the pinned image, removed afterwards
  - only the workspace is mounted (at /workspace, the working dir)
  - env carries exactly one secret: the model API key
  - entrypoint applies a default-deny egress firewall (allowlist = model API);
    a firewall failure aborts the attempt (fail-safe) rather than running open
  - hard wall-clock kill takes down the whole container

The agent's stdout (its machine-readable transcript) streams to
attempt_dir/transcript.jsonl; stderr to attempt_dir/agent_stderr.log.
"""
import json
import subprocess
import time
import uuid
from dataclasses import dataclass
from pathlib import Path

from .config import (CODEX_CATALOG_SOURCE, CODEX_CATALOG_TARGET, RELAY_SOURCE, RELAY_TARGET,
                     RunConfig, uses_codex_catalog)


@dataclass
class SandboxResult:
    exit_code: int | None  # None if we killed it
    duration_seconds: float
    timed_out: bool
    transcript_path: Path
    stderr_path: Path
    infra_error: str | None = None  # set when the sandbox itself failed
    provider_wait_seconds: float = 0.0  # relay 429/quota waits not charged to the wall clock


class ProviderWaits:
    """Running total of the relay's upstream_retries delays (seconds spent waiting out
    provider 429 / quota refusals), read incrementally from its trajectory.jsonl.
    Only complete lines are read, so a line still being written is picked up later."""

    def __init__(self, path: Path):
        self.path, self.pos, self.total = path, 0, 0.0

    def seconds(self) -> float:
        try:
            with open(self.path, "rb") as f:
                f.seek(self.pos)
                data = f.read()
        except OSError:
            return self.total
        end = data.rfind(b"\n")
        if end < 0:
            return self.total
        self.pos += end + 1
        for line in data[:end + 1].splitlines():
            try:
                entry = json.loads(line)
            except ValueError:
                continue
            for retry in entry.get("upstream_retries") or []:
                self.total += float(retry.get("delay_s") or 0)
        return self.total


def run_in_sandbox(cfg: RunConfig, agent_cmd: list, agent_env: dict,
                   workspace: Path, attempt_dir: Path) -> SandboxResult:
    transcript = attempt_dir / "transcript.jsonl"
    stderr_log = attempt_dir / "agent_stderr.log"
    prompt_file = workspace / "PROMPT.md"

    if cfg.sandbox.mode == "docker":
        name = f"coding-agent-{uuid.uuid4().hex[:12]}"
        cmd = [
            "docker", "run", "--rm", "-i",
            "--name", name,
            "-v", f"{workspace.resolve()}:/workspace",
            "-w", "/workspace",
            "--cap-add", "NET_ADMIN",  # entrypoint needs it to install the firewall, then drops to the agent user
            "--memory", cfg.sandbox.memory,
            "--cpus", str(cfg.sandbox.cpus),
            "--pids-limit", "1024",
            "-e", f"ALLOWED_DOMAINS={','.join(cfg.allowed_domains)}",
        ]
        if cfg.record_trajectory:
            if not RELAY_SOURCE.is_file():  # docker would mount an empty DIRECTORY in its place
                return SandboxResult(None, 0.0, False, transcript, stderr_log,
                                     infra_error=f"Sandbox launch failed: relay source missing: {RELAY_SOURCE}")
            traj_dir = attempt_dir / "trajectory"
            traj_dir.mkdir(exist_ok=True)
            upstream = ("https://api.anthropic.com" if cfg.agent.cli == "claude"
                        else "https://api.openai.com")
            cmd += ["-v", f"{traj_dir.resolve()}:/trajectory",
                    "-v", f"{RELAY_SOURCE}:{RELAY_TARGET}:ro",  # repo relay over the baked one (see config.RELAY_SOURCE)
                    "-e", f"TRAJ_UPSTREAM={upstream}"]
            if cfg.agent.cli == "claude":  # codex is routed via -c provider flags instead
                agent_env = {**agent_env, "ANTHROPIC_BASE_URL": "http://127.0.0.1:9877"}
        if uses_codex_catalog(cfg.agent):
            if not CODEX_CATALOG_SOURCE.is_file():  # codex would refuse to start: no trial burned
                return SandboxResult(None, 0.0, False, transcript, stderr_log,
                                     infra_error=f"Sandbox launch failed: codex model catalog missing: {CODEX_CATALOG_SOURCE}")
            cmd += ["-v", f"{CODEX_CATALOG_SOURCE}:{CODEX_CATALOG_TARGET}:ro"]  # see config.CODEX_CATALOG_SOURCE
        for key, value in agent_env.items():
            cmd += ["-e", f"{key}={value}"]
        cmd += [cfg.sandbox.image] + agent_cmd
        env = None  # secrets go into the container only via -e, not the docker client env
        kill = lambda: subprocess.run(["docker", "kill", name], capture_output=True)
    elif cfg.sandbox.mode == "host":
        print("⚠️  sandbox.mode=host — UNSANDBOXED run (dev/rung-0 only). "
              "No isolation, no network restrictions.")
        cmd = agent_cmd
        import os
        env = {**os.environ, **agent_env}
        kill = None  # process-group kill handled below
        name = None
    else:
        raise ValueError(cfg.sandbox.mode)

    start = time.monotonic()
    try:
        with open(prompt_file, "rb") as stdin, \
             open(transcript, "wb") as out, open(stderr_log, "wb") as err:
            proc = subprocess.Popen(
                cmd, stdin=stdin, stdout=out, stderr=err, env=env,
                cwd=None if cfg.sandbox.mode == "docker" else str(workspace),
                start_new_session=(cfg.sandbox.mode == "host"),
            )
            waits = (ProviderWaits(attempt_dir / "trajectory" / "trajectory.jsonl")
                     if cfg.limits.exclude_provider_waits and cfg.record_trajectory
                     and cfg.sandbox.mode == "docker" else None)
            try:
                if waits is None:
                    proc.wait(timeout=cfg.limits.wall_clock_seconds)
                else:
                    while True:  # the deadline moves out by every provider wait logged so far
                        left = cfg.limits.wall_clock_seconds + waits.seconds() - (time.monotonic() - start)
                        if left <= 0:
                            raise subprocess.TimeoutExpired(cmd, cfg.limits.wall_clock_seconds)
                        try:
                            proc.wait(timeout=min(left, 30))
                            break
                        except subprocess.TimeoutExpired:
                            continue
                timed_out = False
            except subprocess.TimeoutExpired:
                timed_out = True
                if kill:
                    kill()
                else:
                    import os, signal
                    os.killpg(proc.pid, signal.SIGKILL)
                proc.wait(timeout=60)
    except FileNotFoundError as e:
        # docker (or the agent CLI in host mode) is not installed
        return SandboxResult(None, time.monotonic() - start, False,
                             transcript, stderr_log,
                             infra_error=f"Sandbox launch failed: {e}")

    duration = time.monotonic() - start
    exit_code = None if timed_out else proc.returncode

    # Docker's own failures (bad image, daemon down, firewall abort) surface as
    # exit codes 125-127 from `docker run`, or our entrypoint's dedicated 97.
    infra_error = None
    if cfg.sandbox.mode == "docker" and exit_code not in (0, None):
        tail = ""
        try:
            tail = stderr_log.read_text(errors="replace")[-2000:]
        except OSError:
            pass
        if exit_code in (97, 125, 126, 127):
            infra_error = f"Container/infra failure (exit {exit_code}): {tail}"
        elif "docker" in tail.lower() and ("daemon" in tail.lower() or "docker.sock" in tail.lower()):
            infra_error = f"Docker daemon unreachable (is Docker Desktop running?): {tail[:300]}"

    return SandboxResult(exit_code, duration, timed_out, transcript, stderr_log,
                         infra_error=infra_error,
                         provider_wait_seconds=waits.seconds() if waits is not None else 0.0)
