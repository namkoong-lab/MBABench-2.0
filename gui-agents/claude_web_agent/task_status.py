"""
Task status definitions for the Claude Web Agent retry pipeline.

Provides a two-tier classification:
- AGENT statuses: the agent was invoked and produced an outcome (counts toward agent_attempts)
- PIPELINE statuses: pre-agent infrastructure failures (don't count toward agent_attempts)

In the GUI agent context, download failures and corrupted files are the agent's fault
(unlike TabAI where the download layer is separate infrastructure).
"""

from enum import Enum


class TaskStatus(str, Enum):
    """Status of a task attempt — classifies both agent and pipeline outcomes."""

    # Agent statuses (agent was invoked, count toward agent_attempts)
    SUCCESS = "success"
    TIMEOUT = "timeout"
    PROMPT_FAILED = "prompt_failed"
    DOWNLOAD_FAILED = "download_failed"
    FILE_CORRUPTED = "file_corrupted"
    # Pipeline statuses (pre-agent infra failures, DON'T count)
    NAVIGATION_FAILED = "navigation_failed"
    AUTH_FAILED = "auth_failed"
    UPLOAD_FAILED = "upload_failed"
    RATE_LIMITED = "rate_limited"
    UNKNOWN = "unknown"


AGENT_STATUSES = {
    TaskStatus.SUCCESS,
    TaskStatus.TIMEOUT,
    TaskStatus.PROMPT_FAILED,
    TaskStatus.DOWNLOAD_FAILED,
    TaskStatus.FILE_CORRUPTED,
}

PIPELINE_STATUSES = {
    TaskStatus.NAVIGATION_FAILED,
    TaskStatus.AUTH_FAILED,
    TaskStatus.UPLOAD_FAILED,
    TaskStatus.RATE_LIMITED,
    TaskStatus.UNKNOWN,
}


class PipelineError(Exception):
    """Raised for pre-agent infrastructure failures (browser, nav, auth, file upload)."""

    def __init__(self, status: TaskStatus, message: str = ""):
        self.status = status
        super().__init__(message or status.value)
