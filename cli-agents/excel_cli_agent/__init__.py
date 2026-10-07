"""
Excel CLI Agent Package
Intelligent Excel automation with iterative task execution
"""

from .mcp_client import ExcelMCPClient
from .task_executor import ExcelTaskExecutor, TaskStatus, ExecutionStep, TaskExecution
from .cli import ExcelCLIAgent

__version__ = "0.1.0"

__all__ = [
    "ExcelMCPClient",
    "ExcelTaskExecutor",
    "TaskStatus",
    "ExecutionStep",
    "TaskExecution",
    "ExcelCLIAgent",
]