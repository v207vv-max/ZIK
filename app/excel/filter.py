from __future__ import annotations

from .reader import ExcelRow


class ResultFilter:
    """
    Фильтр по колонке 'Результат'.

    Значения берутся непосредственно из Excel.
    ZIK не ограничивает их заранее заданным списком.

    Например Excel может содержать:

        Пусто
        boradi
        ko'tarmadi
        ko'chada
        ARMADA
        BORODA
        KOCHADAN
        BUSY
        NO_ANSWER

    Все эти значения автоматически появятся в фильтре.
    """

    RESULT_COLUMN = "Результат"

    def __init__(
        self,
        rows: list[ExcelRow],
    ) -> None:
        self.rows = rows

        # None = фильтр выключен.
        #
        # set(...) = показывать только выбранные значения.
        self.selected: set[str] | None = None

    # =========================================================
    # NORMALIZE
    # =========================================================

    @staticmethod
    def normalize(
        value: object,
    ) -> str:
        """
        Приводит значение Excel к отображаемому виду.

        None / пустая строка -> 'Пусто'
        """

        if value is None:
            return "Пусто"

        value = str(value).strip()

        if not value:
            return "Пусто"

        return value

    # =========================================================
    # VALUES
    # =========================================================

    def get_values(self) -> list[str]:
        """
        Возвращает уникальные значения Результат,
        реально существующие в текущем Excel.

        'Пусто' всегда находится первым.
        """

        values: list[str] = []
        seen: set[str] = set()

        # -----------------------------------------------------
        # Сначала пустые значения.
        # -----------------------------------------------------

        has_empty = False

        for row in self.rows:
            value = self.normalize(row.result)

            if value == "Пусто":
                has_empty = True
                break

        if has_empty:
            values.append("Пусто")
            seen.add("Пусто")

        # -----------------------------------------------------
        # Остальные реальные значения Excel.
        # -----------------------------------------------------

        for row in self.rows:
            value = self.normalize(row.result)

            if value in seen:
                continue

            seen.add(value)
            values.append(value)

        return values

    # =========================================================
    # APPLY
    # =========================================================

    def apply(
        self,
        rows: list[ExcelRow] | None = None,
    ) -> list[ExcelRow]:
        """
        Применяет фильтр.

        ВАЖНО:
        ничего в Excel не удаляет.

        Просто возвращает список строк,
        которые должны отображаться в ZIK.
        """

        source = (
            self.rows
            if rows is None
            else rows
        )

        # Фильтр выключен.
        if self.selected is None:
            return list(source)

        # Показываем только выбранные значения.
        return [
            row
            for row in source
            if self.normalize(row.result)
            in self.selected
        ]

    # =========================================================
    # SELECTION
    # =========================================================

    def set_selected(
        self,
        values: set[str] | None,
    ) -> None:
        """
        Устанавливает выбранные значения.

        None = показать всё.
        """

        if values is None:
            self.selected = None
            return

        self.selected = set(values)

    def clear(self) -> None:
        """
        Полностью отключает фильтр.
        """

        self.selected = None

    # =========================================================
    # UPDATE
    # =========================================================

    def update_rows(
        self,
        rows: list[ExcelRow],
    ) -> None:
        """
        Обновляет данные после повторного чтения Excel.
        """

        self.rows = rows

        # Если выбранные значения больше не существуют
        # в новом Excel — оставляем только существующие.
        if self.selected is not None:
            available = set(
                self.get_values()
            )

            self.selected &= available

            if not self.selected:
                self.selected = None

    # =========================================================
    # STATUS
    # =========================================================

    @property
    def is_active(self) -> bool:
        return self.selected is not None