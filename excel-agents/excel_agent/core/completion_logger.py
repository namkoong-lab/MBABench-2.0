"""
Completion Logger - Slimmed-down logging for task and prompt timing.
"""

import json
import logging
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)


class CompletionLogger:
    """Lightweight logger for tracking task and prompt completion times."""

    def __init__(
        self,
        log_dir="json_logs",
        task_identifier=None,
        agent_name=None,
        prompt_version=None,
        task_source=None,
        max_sec_per_task=None,
    ):
        """Initialize completion logger.

        Args:
            log_dir: Directory for JSON completion files. Caller should pass
                     the {YYYYMMDD}_{agentLabel}/json_logs/ path so files are
                     written directly to the run directory.
        """
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        # Create session log file with task identifier and agent if provided
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if task_identifier:
            # Clean task identifier for filename (remove special chars)
            clean_id = "".join(c for c in task_identifier if c.isalnum() or c in "._-")[
                :50
            ]
            if agent_name:
                self.session_file = (
                    self.log_dir
                    / f"completion_{agent_name}_{timestamp}_{clean_id}.json"
                )
            else:
                self.session_file = (
                    self.log_dir / f"completion_{timestamp}_{clean_id}.json"
                )
        else:
            if agent_name:
                self.session_file = (
                    self.log_dir / f"completion_{agent_name}_session_{timestamp}.json"
                )
            else:
                self.session_file = (
                    self.log_dir / f"completion_session_{timestamp}.json"
                )

        # Session data structure
        self.session_data = {
            "session_start": datetime.now().isoformat(),
            "agent_name": agent_name,
            "prompt_version": prompt_version,
            "task_source": task_source,
            "max_sec_per_task": max_sec_per_task,
            "tasks": [],
        }

        # Current task tracking
        self.current_task = None
        self.current_prompt = None

    def start_task(self, task_name, attempt_number=None):
        """Log start of a task. This is when the JSON file first touches disk."""
        task_data = {
            "task_name": task_name,
            "attempt_number": attempt_number,
            "start_time": datetime.now().isoformat(),
            "prompts": [],
        }

        self.current_task = task_data
        self.session_data["tasks"].append(task_data)

        # Write to file immediately — this is the first disk write
        self._write_session()

        logger.info("Task started: %s", task_name)

    def end_task(self, task_status=None, error_msg=None):
        """Log end of current task.

        Args:
            task_status: TaskStatus enum (or string). Determines agent_failed fields.
            error_msg: Optional error message
        """
        if not self.current_task:
            return

        from .file_organizer import AGENT_STATUSES, TaskStatus

        self.current_task["end_time"] = datetime.now().isoformat()

        # Resolve task_status string
        if task_status is not None:
            status_str = (
                task_status.value if hasattr(task_status, "value") else str(task_status)
            )
        else:
            status_str = "unknown"
        self.current_task["task_status"] = status_str

        # Resolve the TaskStatus enum for classification
        try:
            status_enum = TaskStatus(status_str)
        except ValueError:
            status_enum = TaskStatus.UNKNOWN

        # Derive agent_failed / agent_failed_reason
        if status_enum == TaskStatus.SUCCESS:
            self.current_task["agent_failed"] = False
            self.current_task["agent_failed_reason"] = None
        elif status_enum in AGENT_STATUSES:
            # Agent ran but failed (timeout, prompt_failed)
            self.current_task["agent_failed"] = True
            self.current_task["agent_failed_reason"] = status_str
        else:
            # Pipeline failure (download_failed, file_corrupted, etc.)
            self.current_task["agent_failed"] = None
            self.current_task["agent_failed_reason"] = None

        # deprecated: False for agent attempts (post-processing may upgrade to True),
        #             None for pipeline failures (not applicable)
        if status_enum in AGENT_STATUSES:
            self.current_task["deprecated"] = False
            self.current_task["deprecated_reason"] = None
        else:
            self.current_task["deprecated"] = None
            self.current_task["deprecated_reason"] = None

        if error_msg:
            self.current_task["error"] = error_msg

        # Calculate duration
        start_time = datetime.fromisoformat(self.current_task["start_time"])
        end_time = datetime.fromisoformat(self.current_task["end_time"])
        duration = (end_time - start_time).total_seconds()
        self.current_task["duration_seconds"] = duration

        logger.info(
            "Task completed: %s (%.1fs)", self.current_task["task_name"], duration
        )

        # Write to file
        self._write_session()
        self.current_task = None

    def set_solution_file(self, path):
        """Record the downloaded solution workbook's absolute path.

        The runner reads this back from the completion JSON, so the file
        that gets uploaded is the file the engine actually downloaded —
        never a directory-glob guess."""
        self.session_data["solution_file"] = str(path) if path else None
        if self.session_data["tasks"]:
            self.session_data["tasks"][-1]["solution_file"] = (
                str(path) if path else None
            )
        self._write_session()

    def update_task_status(self, task_status, validation_error=None):
        """Update the last ended task's status in-place.

        Args:
            task_status: TaskStatus enum (or string)
            validation_error: Optional validation error message to include
        """
        from .file_organizer import AGENT_STATUSES, TaskStatus

        if not self.session_data["tasks"]:
            return

        last_task = self.session_data["tasks"][-1]
        status_str = (
            task_status.value if hasattr(task_status, "value") else str(task_status)
        )
        last_task["task_status"] = status_str

        # Resolve the TaskStatus enum for classification
        try:
            status_enum = TaskStatus(status_str)
        except ValueError:
            status_enum = TaskStatus.UNKNOWN

        # Derive agent_failed / agent_failed_reason
        if status_enum == TaskStatus.SUCCESS:
            last_task["agent_failed"] = False
            last_task["agent_failed_reason"] = None
        elif status_enum in AGENT_STATUSES:
            last_task["agent_failed"] = True
            last_task["agent_failed_reason"] = status_str
        else:
            last_task["agent_failed"] = None
            last_task["agent_failed_reason"] = None

        if validation_error:
            last_task["validation_error"] = str(validation_error)

        self._write_session()

    def start_prompt(self, prompt_text):
        """Log start of a prompt (before pressing enter)."""
        if not self.current_task:
            return

        prompt_data = {
            "prompt_text": prompt_text,
            "start_time": datetime.now().isoformat(),
        }

        self.current_prompt = prompt_data
        self.current_task["prompts"].append(prompt_data)

        display = prompt_text if len(prompt_text) <= 100 else prompt_text[:100] + "..."
        logger.info("Prompt started: %s", display)

    def end_prompt(self, success=True, response_length=None):
        """Log end of current prompt (after receiving confirmation)."""
        if not self.current_prompt:
            return

        self.current_prompt["end_time"] = datetime.now().isoformat()
        self.current_prompt["success"] = success
        if response_length:
            self.current_prompt["response_length"] = response_length

        # Calculate duration
        start_time = datetime.fromisoformat(self.current_prompt["start_time"])
        end_time = datetime.fromisoformat(self.current_prompt["end_time"])
        duration = (end_time - start_time).total_seconds()
        self.current_prompt["duration_seconds"] = duration

        logger.info("Prompt completed: %.1fs", duration)

        # Write to file
        self._write_session()
        self.current_prompt = None

    def _write_session(self):
        """Write session data to file."""
        try:
            with open(self.session_file, "w") as f:
                json.dump(self.session_data, f, indent=2)
        except Exception as e:
            logger.warning("Could not write completion log: %s", e)

    def get_session_summary(self):
        """Get summary of current session."""
        if not self.session_data["tasks"]:
            return "No tasks completed yet."

        total_tasks = len(self.session_data["tasks"])
        successful_tasks = sum(
            1
            for task in self.session_data["tasks"]
            if task.get("agent_failed") is False
        )

        total_prompts = sum(
            len(task.get("prompts", [])) for task in self.session_data["tasks"]
        )
        successful_prompts = sum(
            len([p for p in task.get("prompts", []) if p.get("success", False)])
            for task in self.session_data["tasks"]
        )

        return {
            "total_tasks": total_tasks,
            "successful_tasks": successful_tasks,
            "total_prompts": total_prompts,
            "successful_prompts": successful_prompts,
            "session_file": str(self.session_file),
        }
