from __future__ import annotations

from .models import CallResult


ARMADA_LIMIT_SECONDS = 25.0


def classify_call_duration(duration: float) -> CallResult:
    """
    Classify a completed answered call.

    < 25 sec  -> ARMADA
    >= 25 sec -> BORODA
    """

    if duration < ARMADA_LIMIT_SECONDS:
        return CallResult.ARMADA

    return CallResult.BORODA