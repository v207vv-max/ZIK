from __future__ import annotations

import logging
import threading
import time
from pathlib import Path

import openpyxl


logger = logging.getLogger("ZIK.ExcelWriter")


class ExcelWriter:
    """
    Direct Excel writer for ZIK.

    Current Sverka.xlsx layout:

        A = Дата
        B = Продукт
        C = Имя
        D = Тип Оплаты
        E = Номер
        F = status/code
        G = Результат

    ZIK writes ONLY the final call result to column G.

    This class does not use WPS, does not move a WPS cursor,
    and does not send keyboard/mouse commands.
    """

    RESULT_COLUMN = "Результат"
    RESULT_COLUMN_INDEX = 7

    # Excel/WPS may briefly keep the file locked while closing/saving.
    SAVE_RETRIES = 5
    SAVE_RETRY_DELAY_SECONDS = 0.5

    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self._lock = threading.RLock()

    # =========================================================
    # OPEN
    # =========================================================

    def _open(self):
        if not self.path.exists():
            raise FileNotFoundError(
                f"Excel-файл не найден: {self.path}"
            )

        return openpyxl.load_workbook(
            self.path,
            data_only=False,
        )

    # =========================================================
    # RESULT COLUMN
    # =========================================================

    def _find_result_column(self, worksheet) -> int:
        """
        Finds the existing Результат column.

        The current workbook has Результат in G.

        Important:
        F is the status/code column and must NEVER be used
        as the result column.
        """

        # First prefer the actual header.
        for cell in worksheet[1]:
            value = (
                str(cell.value).strip()
                if cell.value is not None
                else ""
            )

            if value == self.RESULT_COLUMN:
                return cell.column

        # If the header is empty/corrupted, the known workbook
        # layout still puts Result in G.
        if worksheet.max_column >= self.RESULT_COLUMN_INDEX:
            return self.RESULT_COLUMN_INDEX

        # If G does not exist yet, create exactly G.
        # Never create a new column after G.
        return self.RESULT_COLUMN_INDEX

    def ensure_result_column(self, worksheet) -> int:
        """
        Ensures that G1 contains 'Результат' and returns column 7.

        Column F is never changed.
        """

        result_column = self._find_result_column(worksheet)

        if result_column != self.RESULT_COLUMN_INDEX:
            raise ValueError(
                "Некорректная структура Excel: "
                f"колонка '{self.RESULT_COLUMN}' найдена "
                f"в колонке {result_column}, ожидалась G."
            )

        header = worksheet.cell(
            row=1,
            column=self.RESULT_COLUMN_INDEX,
        )

        if (
            header.value is None
            or not str(header.value).strip()
        ):
            header.value = self.RESULT_COLUMN

            logger.info(
                "EXCEL: восстановлен заголовок '%s' в G1",
                self.RESULT_COLUMN,
            )

        return self.RESULT_COLUMN_INDEX

    # =========================================================
    # WRITE
    # =========================================================

    def write_result(
        self,
        excel_row: int,
        result: str,
    ) -> None:
        """
        Writes a final call result to G[excel_row].

        Example:

            write_result(185, "BORODA")

        produces:

            G185 = BORODA

        F185 is left untouched.
        """

        if excel_row < 2:
            raise ValueError(
                f"Некорректная строка Excel: {excel_row}"
            )

        result_text = (
            str(result).strip()
            if result is not None
            else ""
        )

        if not result_text:
            raise ValueError(
                "Нельзя записать пустой результат звонка."
            )

        with self._lock:
            last_error: Exception | None = None

            for attempt in range(
                1,
                self.SAVE_RETRIES + 1,
            ):
                workbook = None

                try:
                    workbook = self._open()
                    worksheet = workbook.active

                    result_column = self.ensure_result_column(
                        worksheet
                    )

                    if excel_row > worksheet.max_row:
                        raise ValueError(
                            "Строка Excel выходит за пределы "
                            f"таблицы: {excel_row} > {worksheet.max_row}"
                        )

                    # Write ONLY to G.
                    worksheet.cell(
                        row=excel_row,
                        column=result_column,
                        value=result_text,
                    )

                    workbook.save(self.path)

                    # Reopen the saved workbook and verify the value.
                    # This catches cases where a save completed but the
                    # expected cell was not actually persisted.
                    workbook.close()
                    workbook = None

                    verify_workbook = None

                    try:
                        verify_workbook = openpyxl.load_workbook(
                            self.path,
                            data_only=False,
                            read_only=True,
                        )

                        verify_sheet = verify_workbook.active
                        saved_value = verify_sheet.cell(
                            row=excel_row,
                            column=self.RESULT_COLUMN_INDEX,
                        ).value

                    finally:
                        if verify_workbook is not None:
                            verify_workbook.close()

                    if str(saved_value).strip() != result_text:
                        raise IOError(
                            "Проверка записи Excel не пройдена: "
                            f"G{excel_row} содержит {saved_value!r}, "
                            f"ожидалось {result_text!r}"
                        )

                    logger.info(
                        "EXCEL: результат записан | "
                        "row=%s | column=G | result=%s",
                        excel_row,
                        result_text,
                    )

                    return

                except (
                    PermissionError,
                    OSError,
                    IOError,
                ) as exc:
                    last_error = exc

                    logger.warning(
                        "EXCEL: файл занят или не удалось сохранить "
                        "| попытка=%s/%s | row=%s | error=%s",
                        attempt,
                        self.SAVE_RETRIES,
                        excel_row,
                        exc,
                    )

                    if attempt < self.SAVE_RETRIES:
                        time.sleep(
                            self.SAVE_RETRY_DELAY_SECONDS
                        )

                except Exception:
                    logger.exception(
                        "EXCEL: ошибка записи "
                        "| row=%s | result=%s",
                        excel_row,
                        result_text,
                    )
                    raise

                finally:
                    if workbook is not None:
                        workbook.close()

            raise PermissionError(
                "Не удалось сохранить Excel-файл после "
                f"{self.SAVE_RETRIES} попыток. "
                "Закройте Sverka.xlsx в WPS/Excel и повторите звонок."
            ) from last_error
