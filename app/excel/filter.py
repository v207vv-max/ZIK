from __future__ import annotations

from collections.abc import Iterable
from typing import TypeVar

from .reader import ExcelRow


T = TypeVar("T")


class ColumnFilter:
    """Normalization helpers used by ZIK checkbox filters."""

    EMPTY_LABEL = "Пусто"

    @classmethod
    def normalize(cls, value: object) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @classmethod
    def display_value(cls, value: object) -> str:
        normalized = cls.normalize(value)
        return normalized if normalized else cls.EMPTY_LABEL

    @classmethod
    def to_stored(cls, display_value: str) -> str:
        return "" if display_value == cls.EMPTY_LABEL else display_value

    @classmethod
    def apply(
        cls,
        rows: Iterable[T],
        selected: set[str] | None,
        getter,
    ) -> list[T]:
        rows = list(rows)

        if selected is None:
            return rows

        return [
            row
            for row in rows
            if cls.normalize(getter(row)) in selected
        ]


class ResultFilter:
    """
    Checkbox filter for Excel column G: «Результат».

    None means no filter: all result values are allowed.
    An empty selection is also treated as no filter so that the UI
    can never accidentally leave the application in a broken state.
    """

    RESULT_COLUMN = "Результат"

    def __init__(self, rows: list[ExcelRow]) -> None:
        self.rows = rows
        self._selected: set[str] | None = None

    @staticmethod
    def normalize(value: object) -> str:
        normalized = ColumnFilter.normalize(value)
        return (
            normalized
            if normalized
            else ColumnFilter.EMPTY_LABEL
        )

    @property
    def selected(self) -> set[str] | None:
        if self._selected is None:
            return None
        return set(self._selected)

    @property
    def is_active(self) -> bool:
        return self._selected is not None

    def get_values(self) -> list[str]:
        values = {
            self.normalize(row.result)
            for row in self.rows
        }

        return sorted(
            values,
            key=lambda value: (
                value == ColumnFilter.EMPTY_LABEL,
                value.casefold(),
            ),
        )

    def apply(
        self,
        rows: list[ExcelRow],
    ) -> list[ExcelRow]:
        if self._selected is None:
            return list(rows)

        return [
            row
            for row in rows
            if self.normalize(row.result)
            in self._selected
        ]

    def set_selected(
        self,
        selected: set[str] | None,
    ) -> None:
        if selected is None:
            self._selected = None
            return

        normalized = {
            self.normalize(value)
            for value in selected
        }

        all_values = set(self.get_values())

        if not normalized or normalized == all_values:
            self._selected = None
            return

        self._selected = normalized

    def clear(self) -> None:
        self._selected = None

    def update_rows(
        self,
        rows: list[ExcelRow],
    ) -> None:
        self.rows = rows

        if self._selected is None:
            return

        available = set(self.get_values())
        self._selected &= available

        if not self._selected or self._selected == available:
            self._selected = None
