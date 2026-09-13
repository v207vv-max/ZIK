from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from datetime import datetime


class CallState(str, Enum):
    IDLE = "IDLE"
    OUTGOING = "OUTGOING"
    IN_CALL = "IN_CALL"
    BUSY = "BUSY"
    ENDED = "ENDED"
    FAILED = "FAILED"


class CallResult(str, Enum):
    UNKNOWN = "UNKNOWN"
    ARMADA = "ARMADA"
    BORODA = "BORODA"
    KOCHADAN = "KOCHADAN"
    BUSY = "BUSY"
    NO_ANSWER = "NO_ANSWER"
    ERROR = "ERROR"


@dataclass
class Call:
    id: int | None
    phone: str
    state: CallState
    result: CallResult

    outgoing_at: datetime | None = None
    started_at: datetime | None = None
    ended_at: datetime | None = None

    manual_result: CallResult | None = None

    @property
    def ring_duration(self) -> float | None:
        if not self.outgoing_at:
            return None

        end = self.started_at or self.ended_at

        if not end:
            return None

        return (end - self.outgoing_at).total_seconds()

    @property
    def call_duration(self) -> float | None:
        if not self.started_at or not self.ended_at:
            return None

        return (self.ended_at - self.started_at).total_seconds()