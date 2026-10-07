"""
CompletionLogger for the Web Agent Engine.

Writes JSON completion logs to disk on every state change for crash safety.

JSON schema:
{
    "session_start": "ISO-8601",
    "agent_name": "claude_web",
    "prompt_version": 1,
    "task_source": "my_tasks",
    "tasks": [{
        "task_name": "Example-Task",
        "attempt_number": 1,
        "start_time": "ISO-8601",
        "end_time": "ISO-8601",
        "task_status": "success",
        "agent_failed": false,
        "agent_failed_reason": null,
        "deprecated": false,
        "deprecation_reason": null,
        "duration_seconds": 1200.0,
        "prompts": [{
            "prompt_text": "...",
            "start_time": "ISO-8601",
            "end_time": "ISO-8601",
            "success": true,
            "duration_seconds": 300.0
        }]
    }]
}
"""

import json
import logging
import re
from datetime import datetime
from pathlib import Path

from claude_web_agent.task_status import TaskStatus

logger = logging.getLogger(__name__)


class CompletionLogger:
    """
    Tracks task and prompt timing, writes JSON completion logs.

    Usage:
        cl = CompletionLogger(log_dir, "task-name", "claude_web", 1, "my_tasks")
        cl.start_task("task-name", attempt_number=1)
        cl.start_prompt("prompt text...")
        cl.end_prompt(success=True, response_length=1234)
        cl.end_task(TaskStatus.SUCCESS)
    """

    def __init__(
        self,
        log_dir: str | Path = "json_logs",
        task_identifier: str | None = None,
        agent_name: str = "claude_web",
        prompt_version: int = 1,
        task_source: str = "my_tasks",
    ):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        self.task_identifier = task_identifier
        self.agent_name = agent_name
        self.prompt_version = prompt_version
        self.task_source = task_source
        self.session_start = datetime.now()

        self.session_data = {
            "session_start": self.session_start.isoformat(),
            "agent_name": agent_name,
            "prompt_version": prompt_version,
            "task_source": task_source,
            "tasks": [],
        }

        self.current_task = None
        self.current_prompt = None

        # Build session file path
        timestamp = self.session_start.strftime("%Y%m%d_%H%M%S")
        clean_id = self._clean_name(task_identifier or "unknown")
        filename = f"completion_{agent_name}_{timestamp}_{clean_id}.json"
        self.session_file = self.log_dir / filename

        # Write initial state
        self._write_to_disk()

    @staticmethod
    def _clean_name(name: str) -> str:
        """Sanitize a name for use in filenames."""
        name = name.replace("/", "-").replace("\\", "-").replace(" ", "_")
        name = re.sub(r"[^a-zA-Z0-9._-]", "", name)
        return name

    def _write_to_disk(self):
        """Persist current session data to disk (crash safety)."""
        try:
            with open(self.session_file, "w") as f:
                json.dump(self.session_data, f, indent=2)
        except Exception as e:
            logger.error(f"Failed to write completion log: {e}")

    def start_task(self, task_name: str, attempt_number: int = 1):
        """Begin tracking a new task attempt."""
        self.current_task = {
            "task_name": task_name,
            "attempt_number": attempt_number,
            "start_time": datetime.now().isoformat(),
            "end_time": None,
            "task_status": None,
            "agent_failed": None,
            "agent_failed_reason": None,
            "deprecated": False,
            "deprecation_reason": None,
            "duration_seconds": None,
            "prompts": [],
        }
        self._write_to_disk()

    def end_task(self, task_status: TaskStatus):
        """
        Finalize the current task with a status.

        Derives agent_failed and agent_failed_reason from the status:
        - SUCCESS → agent_failed=False
        - Any other status → agent_failed=True, reason=status.value
        """
        if not self.current_task:
            logger.warning("end_task called with no active task")
            return

        now = datetime.now()
        start = datetime.fromisoformat(self.current_task["start_time"])
        duration = (now - start).total_seconds()

        self.current_task["end_time"] = now.isoformat()
        self.current_task["task_status"] = task_status.value
        self.current_task["duration_seconds"] = round(duration, 2)

        # Derive agent_failed fields
        if task_status == TaskStatus.SUCCESS:
            self.current_task["agent_failed"] = False
            self.current_task["agent_failed_reason"] = None
        else:
            self.current_task["agent_failed"] = True
            self.current_task["agent_failed_reason"] = task_status.value

        self.session_data["tasks"].append(self.current_task)
        self.current_task = None
        self._write_to_disk()

    def start_prompt(self, prompt_text: str):
        """Begin tracking a prompt within the current task."""
        self.current_prompt = {
            "prompt_text": prompt_text,
            "start_time": datetime.now().isoformat(),
            "end_time": None,
            "success": None,
            "duration_seconds": None,
        }
        self._write_to_disk()

    def end_prompt(self, success: bool, response_length: int = 0):
        """Finalize the current prompt."""
        if not self.current_prompt:
            logger.warning("end_prompt called with no active prompt")
            return

        now = datetime.now()
        start = datetime.fromisoformat(self.current_prompt["start_time"])
        duration = (now - start).total_seconds()

        self.current_prompt["end_time"] = now.isoformat()
        self.current_prompt["success"] = success
        self.current_prompt["duration_seconds"] = round(duration, 2)
        if response_length:
            self.current_prompt["response_length"] = response_length

        if self.current_task:
            self.current_task["prompts"].append(self.current_prompt)

        self.current_prompt = None
        self._write_to_disk()

    def save(self, output_dir: Path | str | None = None) -> Path:
        """
        Save session data to a specific output directory (copies the session file).

        If output_dir differs from self.log_dir, copies to the new location.
        Returns the path to the saved file.
        """
        if output_dir is None:
            return self.session_file

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        dest = output_dir / self.session_file.name
        if dest != self.session_file:
            import shutil

            shutil.copy2(self.session_file, dest)
            logger.info(f"Copied completion log to: {dest}")

        return dest
