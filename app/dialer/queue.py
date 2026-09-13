from __future__ import annotations

from collections import deque
from dataclasses import dataclass


# Максимум 2 попытки:
# 1-я попытка
# 2-я попытка
# после второй неудачи -> следующий номер
MAX_ATTEMPTS = 2


@dataclass
class QueueItem:
    phone: str
    excel_row: int | None = None
    attempts: int = 0

    @property
    def can_retry(self) -> bool:
        return self.attempts < MAX_ATTEMPTS

    def next_attempt(self) -> int:
        self.attempts += 1
        return self.attempts


class CallQueue:
    def __init__(self) -> None:
        self._queue: deque[QueueItem] = deque()
        self._current: QueueItem | None = None

    def add(
        self,
        phone: str,
        excel_row: int | None = None,
    ) -> None:
        phone = phone.strip()

        if not phone:
            raise ValueError("Phone number is empty")

        self._queue.append(
            QueueItem(
                phone=phone,
                excel_row=excel_row,
            )
        )

    def add_many(
        self,
        items: list[tuple[str, int | None]],
    ) -> None:
        for phone, excel_row in items:
            self.add(
                phone,
                excel_row,
            )

    def start_next(self) -> QueueItem | None:
        if self._current is not None:
            raise RuntimeError(
                "Current queue item is still active"
            )

        if not self._queue:
            return None

        self._current = self._queue.popleft()

        return self._current

    @property
    def current(self) -> QueueItem | None:
        return self._current

    def mark_attempt(self) -> int:
        if self._current is None:
            raise RuntimeError(
                "There is no active queue item"
            )

        return self._current.next_attempt()

    def retry_current(self) -> bool:
        if self._current is None:
            raise RuntimeError(
                "There is no active queue item"
            )

        if self._current.can_retry:
            self._queue.appendleft(
                self._current
            )

            self._current = None

            return True

        return False

    def complete_current(self) -> None:
        self._current = None

    def has_items(self) -> bool:
        return bool(self._queue)

    def size(self) -> int:
        return len(self._queue)

    def clear(self) -> None:
        self._queue.clear()
        self._current = None