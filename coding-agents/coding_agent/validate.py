"""Output validation: the runner — never the agent — decides what happened.

Success requires ALL of:
  - workspace/solution.xlsx exists,
  - it loads as a valid workbook (openpyxl),
  - its sha256 differs from EVERY seeded file in the manifest
    (an untouched/renamed input can never be banked as a solution),
  - duration exceeded the junk threshold (else needs_review).

Verdicts and their consequences:
  success        -> recorded, counts as the attempt
  timeout        -> recorded as failed attempt (partial workbook kept)
  agent_failure  -> recorded as failed attempt
  infra_failure  -> NOT recorded in the DB; the agent never got a fair attempt
  needs_review   -> NOT recorded; held locally for a human decision
"""
import json
from dataclasses import dataclass
from pathlib import Path

from .sandbox import SandboxResult
from .workspace import Attempt, sha256_file

# Transcript signatures of provider-side auth/quota problems: the agent never
# got a fair attempt, so these classify as infra, not agent, failures.
INFRA_SIGNATURES = (
    "invalid x-api-key", "authentication_error", "invalid api key",
    "incorrect api key", "credit balance is too low", "insufficient_quota",
    "billing", "401 unauthorized",
    # TensorBlock Forge, 2026-09-24: HTTP 402 "Unable to complete the request. Please
    # check your account balance and limits." when the credit balance cannot cover a call.
    "account balance and limits",
)

# Codex's last words when it gave up on a model call after exhausting its retries on a
# provider/gateway error. The run ended because of the provider, not the agent, even if a
# workbook was already saved - 2026-09-25: a GLM attempt cut off mid-run by TensorBlock 429s
# (row 3981) was banked as success. Infra: no row, the runner retries the task.
RETRY_EXHAUSTED_SIGNATURES = tuple(f"exceeded retry limit, last status: {code}"
                                   for code in ("429", "500", "502", "503", "504"))


@dataclass
class Verdict:
    status: str  # success | timeout | agent_failure | infra_failure | needs_review
    reason: str
    solution_path: Path | None


def _solution_is_valid(path: Path) -> tuple[bool, str]:
    if path.stat().st_size == 0:
        return False, "solution.xlsx is empty"
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        sheet_count = len(wb.sheetnames)
        wb.close()
    except Exception as e:  # noqa: BLE001 — any load failure means invalid
        return False, f"solution.xlsx failed to open: {e!r}"
    if sheet_count == 0:
        return False, "solution.xlsx has no sheets"
    return True, f"valid workbook ({sheet_count} sheets)"


def validate(attempt: Attempt, sandbox: SandboxResult, junk_seconds: int) -> Verdict:
    solution = attempt.workspace / "solution.xlsx"

    if sandbox.infra_error:
        return Verdict("infra_failure", sandbox.infra_error, None)

    transcript_sample = ""
    try:
        transcript_sample = sandbox.transcript_path.read_text(errors="replace")[-20000:].lower()
        transcript_sample += sandbox.stderr_path.read_text(errors="replace")[-5000:].lower()
    except OSError:
        pass
    for signature in RETRY_EXHAUSTED_SIGNATURES:
        if signature in transcript_sample:
            return Verdict("infra_failure", f"Agent cut off by the provider ({signature!r})", None)
    for signature in INFRA_SIGNATURES:
        if signature in transcript_sample and not solution.exists():
            return Verdict("infra_failure", f"Provider auth/quota error in transcript ({signature!r})", None)

    if sandbox.timed_out:
        keep = solution if solution.exists() else None
        return Verdict("timeout",
                       f"Wall-clock cap hit after {sandbox.duration_seconds:.0f}s"
                       + (" (partial workbook kept)" if keep else ""),
                       keep)

    if not solution.exists():
        other_xlsx = [p for p in attempt.workspace.rglob("*.xlsx")
                      if sha256_file(p) not in attempt.manifest.values()]
        if other_xlsx:
            return Verdict("needs_review",
                           f"No solution.xlsx, but new workbook(s) exist: "
                           f"{[str(p.relative_to(attempt.workspace)) for p in other_xlsx]}",
                           None)
        return Verdict("agent_failure",
                       f"Agent exited (code {sandbox.exit_code}) without producing solution.xlsx", None)

    ok, detail = _solution_is_valid(solution)
    if not ok:
        return Verdict("agent_failure", detail, None)

    if sha256_file(solution) in attempt.manifest.values():
        return Verdict("agent_failure",
                       "solution.xlsx is byte-identical to a seeded input file — not new work", None)

    if sandbox.duration_seconds < junk_seconds:
        return Verdict("needs_review",
                       f"Suspiciously fast success ({sandbox.duration_seconds:.0f}s < {junk_seconds}s junk guard)",
                       solution)

    return Verdict("success", detail, solution)


def write_verdict(attempt: Attempt, verdict: Verdict, sandbox: SandboxResult) -> Path:
    path = attempt.attempt_dir / "verdict.json"
    path.write_text(json.dumps({
        "status": verdict.status,
        "reason": verdict.reason,
        "solution": str(verdict.solution_path) if verdict.solution_path else None,
        "exit_code": sandbox.exit_code,
        "duration_seconds": round(sandbox.duration_seconds, 1),
        "timed_out": sandbox.timed_out,
        "provider_wait_seconds": round(getattr(sandbox, "provider_wait_seconds", 0.0), 1),
    }, indent=2))
    return path
