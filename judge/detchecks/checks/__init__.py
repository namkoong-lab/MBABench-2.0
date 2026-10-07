"""Check registry: {rubric number: Check subclass}.

To add a check: create detchecks/checks/cNN.py with a Check subclass and add it to
_CLASSES below.
"""
from __future__ import annotations

from ..errors import GradingError
from .base import Check
from .c22 import C22
from .c29 import C29
from .c47 import C47
from .c49 import C49
from .c50 import C50
from .c51 import C51
from .c61 import C61
from .c62 import C62
from .c65 import C65
from .c66 import C66
from .c69 import C69
from .c70 import C70
from .c73 import C73
from .c74 import C74
from .c77 import C77
from .c80 import C80
from .c87 import C87
from .c92 import C92
from .c93 import C93
from .c94 import C94
from .c95 import C95

_CLASSES = (C22, C29, C47, C49, C50, C51, C61, C62, C65, C66, C69, C70, C73, C74, C77, C80, C87, C92, C93, C94, C95)

REGISTRY: dict[int, type] = {c.number: c for c in _CLASSES}
BY_KEY: dict[str, type] = {c.key: c for c in _CLASSES}


def live_numbers() -> list[int]:
    """Checks whose Python verdict counts at scoring (Check.live).  The others still run and
    record a verdict marked "live": false (handoff ruling 2026-10-04; today only 65 is off)."""
    return sorted(n for n, c in REGISTRY.items() if getattr(c, "live", True))


def get(number: int) -> type:
    try:
        return REGISTRY[int(number)]
    except (KeyError, ValueError):
        raise GradingError(f"no deterministic check registered for #{number} "
                           f"(registered: {sorted(REGISTRY)})") from None
