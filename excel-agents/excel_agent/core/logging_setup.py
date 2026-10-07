"""
Logging configuration for the Excel Agent Engine.

Handles logging setup with console and optional file output.
"""

import io
import logging
import os
import sys
from datetime import datetime
from typing import Optional


def configure_safe_stdout():
    """
    Configure stdout/stderr to handle Unicode encoding errors on Windows.

    On Windows with cp1252 encoding, print() statements with emojis will fail.
    This function wraps stdout/stderr to replace unencodable characters with '?'.
    """
    # Only needed on Windows with non-UTF-8 encoding
    if sys.platform == "win32":
        if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
            sys.stdout = io.TextIOWrapper(
                sys.stdout.buffer, encoding=sys.stdout.encoding, errors="replace"
            )
        if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
            sys.stderr = io.TextIOWrapper(
                sys.stderr.buffer, encoding=sys.stderr.encoding, errors="replace"
            )


class SafeStreamHandler(logging.StreamHandler):
    """A StreamHandler that handles Unicode encoding errors gracefully on Windows."""

    def emit(self, record):
        try:
            msg = self.format(record)
            stream = self.stream
            # On Windows with cp1252, encode with errors='replace' to handle emojis
            try:
                stream.write(msg + self.terminator)
            except UnicodeEncodeError:
                # Replace unencodable characters with '?' and try again
                encoded_msg = msg.encode(
                    stream.encoding or "utf-8", errors="replace"
                ).decode(stream.encoding or "utf-8", errors="replace")
                stream.write(encoded_msg + self.terminator)
            self.flush()
        except RecursionError:
            raise
        except Exception:
            self.handleError(record)


def setup_logging(
    config: dict,
    logger_name: str = __name__,
    task_name: str = None,
    agent_name: str = None,
) -> tuple[logging.Logger, Optional[str]]:
    """
    Setup logging with both console and optional file output.

    Args:
        config: Configuration dictionary with logging settings
        logger_name: Name for the logger
        task_name: Optional task name to include in log filename
        agent_name: Optional agent name for log filename prefix (overrides agent_type)

    Returns:
        Tuple of (configured logger instance, log file path or None)
    """
    # Configure the ROOT logger so all child loggers propagate to it
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Close existing handlers (flush pending writes) before clearing
    for handler in root_logger.handlers[:]:
        handler.close()
    root_logger.handlers.clear()

    # Create formatter
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    # Always add console handler to root (use SafeStreamHandler for Windows compatibility)
    console_handler = SafeStreamHandler()
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # Get agent type to determine which config section to use
    agent_type = config.get(
        "agent_type", "tabai"
    )  # Default to tabai for backward compat

    # Add file handler if configured
    logging_config = config.get(agent_type, {}).get("logging", {})

    if logging_config.get("save_to_file", False):
        # Create log directory if it doesn't exist
        log_dir = logging_config.get("log_directory", "logs")
        os.makedirs(log_dir, exist_ok=True)

        # Create timestamped filename with task name (timestamp first for sorting)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name_prefix = agent_name or agent_type
        if task_name:
            # Sanitize task name for filename
            safe_task_name = task_name.replace("/", "-").replace(" ", "_")
            log_filename = f"{name_prefix}_{timestamp}_{safe_task_name}.log"
        else:
            log_filename = f"{name_prefix}_{timestamp}.log"
        log_path = os.path.join(log_dir, log_filename)

        # Add file handler to root (use UTF-8 encoding for emoji support)
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)

        # Get the named logger to return
        logger = logging.getLogger(logger_name)
        logger.info(f"📝 Logging to file: {log_path}")
        return logger, log_path

    # Return named logger (will propagate to root) with None for log path
    return logging.getLogger(logger_name), None


def setup_basic_logging(
    level: int = logging.INFO,
    format_str: str = "%(asctime)s - %(levelname)s - %(message)s",
    stream=None,
) -> None:
    """
    Setup basic logging with Windows-compatible Unicode handling.

    This is a drop-in replacement for logging.basicConfig() that handles
    emoji characters gracefully on Windows (cp1252 encoding).

    Args:
        level: Logging level (default: INFO)
        format_str: Log format string
        stream: Output stream (default: sys.stderr)
    """
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    # Clear existing handlers
    root_logger.handlers.clear()

    # Create formatter
    formatter = logging.Formatter(format_str)

    # Use SafeStreamHandler for Windows compatibility
    handler = SafeStreamHandler(stream or sys.stderr)
    handler.setFormatter(formatter)
    root_logger.addHandler(handler)
