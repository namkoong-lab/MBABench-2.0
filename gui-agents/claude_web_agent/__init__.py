"""
Web Agent - Browser automation for Claude.ai and ChatGPT web interfaces.

This module provides Playwright-based automation for running tasks through
https://claude.ai or https://chatgpt.com.
"""

from .chatgpt_web_agent import ChatGPTWebAgent
from .claude_web_agent import ClaudeWebAgent
from .browser_manager import WebBrowserManager
from .completion_logger import CompletionLogger
from .file_validator import validate_excel_file
from .task_status import (
    AGENT_STATUSES,
    PIPELINE_STATUSES,
    PipelineError,
    TaskStatus,
)
from .web_agent import WebAgent, WebAgentState

__all__ = [
    "AGENT_STATUSES",
    "ChatGPTWebAgent",
    "ClaudeWebAgent",
    "CompletionLogger",
    "PIPELINE_STATUSES",
    "PipelineError",
    "TaskStatus",
    "WebAgent",
    "WebAgentState",
    "WebBrowserManager",
    "validate_excel_file",
]
