"""This file initializes a global singleton logger instance for the entire application."""

import logging
import os
import threading

DEFAULT_LOG_FORMAT = (
    "%(asctime)s - %(threadName)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s"
)
DEFAULT_LOGGER_NAME = "mbabench_judge"
DEFAULT_LOG_TO_TERMINAL = True

_logger = None
_logger_file_handlers: dict[str, logging.FileHandler] = {}


def _get_logger():
    """Get the project logger, initializing on first access."""
    global _logger
    if _logger is not None:
        return _logger

    _logger = logging.getLogger(DEFAULT_LOGGER_NAME)
    _logger.setLevel(logging.DEBUG)

    if DEFAULT_LOG_TO_TERMINAL:
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(logging.DEBUG)
        stream_handler.setFormatter(logging.Formatter(DEFAULT_LOG_FORMAT))
        _logger.addHandler(stream_handler)

    return _logger


def __getattr__(name):
    """Module-level __getattr__ so `from .logger import logger` works lazily."""
    if name == "logger":
        return _get_logger()
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


class _ThreadFilter(logging.Filter):
    """Pass only records emitted by the thread that registered the handler."""

    def __init__(self, thread_id: int):
        super().__init__()
        self.thread_id = thread_id

    def filter(self, record):
        return record.thread == self.thread_id


def add_log_file(file_path: str, level: int = logging.DEBUG) -> None:
    """Add a file handler for the given path. Skips if already registered.

    A handler registered from a worker thread (grade.py --workers N runs one
    attempt per thread) receives only that thread's records, so each
    attempt's log file stays its own. From the main thread it receives
    everything, as before.
    """
    if file_path in _logger_file_handlers:
        return
    formatter = logging.Formatter(DEFAULT_LOG_FORMAT)
    os.makedirs(os.path.dirname(file_path) or ".", exist_ok=True)
    handler = logging.FileHandler(file_path)
    handler.setLevel(level)
    handler.setFormatter(formatter)
    if threading.current_thread() is not threading.main_thread():
        handler.addFilter(_ThreadFilter(threading.get_ident()))
    _get_logger().addHandler(handler)
    _logger_file_handlers[file_path] = handler


def remove_log_file(file_path: str) -> None:
    """Remove and close the file handler for the given path. No-op if not registered."""
    handler = _logger_file_handlers.pop(file_path, None)
    if handler is None:
        return
    _get_logger().removeHandler(handler)
    handler.close()
