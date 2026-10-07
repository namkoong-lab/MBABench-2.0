"""Type inference utilities for cell values."""
from typing import Any


def _infer_type(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return "number"
    if isinstance(value, (str,)):
        return "string"
    return type(value).__name__
