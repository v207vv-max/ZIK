from __future__ import annotations

import logging
import threading
from pathlib import Path

import pythoncom
import win32com.client


logger = logging.getLogger("ZIK.WPS")


class WPSSync:
    """
    Управление уже открытым WPS Таблицами через COM.

    Возможности:
        - найти открытый WPS;
        - найти нужный XLSX;
        - переместить визуальный курсор;
        - записать результат;
        - сохранить книгу.

    Важно:
        WPS не закрывается программой.
    """

    PROGIDS = (
        "ket.Application",
        "KET.Application",
        "et.Application",
        "ET.Application",
    )

    def __init__(self) -> None:
        self.enabled = True
        self._lock = threading.RLock()

    # =========================================================
    # COM
    # =========================================================

    def _connect(self):
        """
        Подключиться к уже запущенному WPS Таблицам.

        Новый WPS процесс намеренно НЕ запускаем.
        Нам нужен именно тот WPS, который уже открыт
        пользователем.
        """

        last_error = None

        for progid in self.PROGIDS:
            try:
                app = win32com.client.GetActiveObject(
                    progid
                )

                logger.debug(
                    "WPS: connected | progid=%s",
                    progid,
                )

                return app

            except Exception as exc:
                last_error = exc

        raise RuntimeError(
            "Не удалось подключиться к открытому WPS Таблицам. "
            "Убедись, что Sverka.xlsx открыт в WPS."
        ) from last_error

    # =========================================================
    # WORKBOOK
    # =========================================================

    @staticmethod
    def _normalize_path(value: str) -> str:
        return str(
            Path(value)
            .resolve()
        ).lower()

    def _find_workbook(
        self,
        app,
        excel_path: Path,
    ):
        target = self._normalize_path(
            str(excel_path)
        )

        workbooks = app.Workbooks

        for index in range(
            1,
            workbooks.Count + 1,
        ):
            workbook = workbooks.Item(index)

            try:
                full_name = str(
                    workbook.FullName
                )

                if (
                    self._normalize_path(full_name)
                    == target
                ):
                    return workbook

            except Exception:
                continue

        raise RuntimeError(
            f"WPS не открыл файл:\n{excel_path}"
        )

    # =========================================================
    # SHEET
    # =========================================================

    @staticmethod
    def _get_first_sheet(workbook):
        return workbook.Worksheets.Item(1)

    # =========================================================
    # RESULT COLUMN
    # =========================================================

    @staticmethod
    def _find_result_column(sheet) -> int:
        """
        Ищем колонку 'Результат' в первой строке.
        Если её нет — создаём.
        """

        used_range = sheet.UsedRange

        max_column = int(
            used_range.Columns.Count
        )

        for column in range(
            1,
            max_column + 1,
        ):
            value = sheet.Cells(
                1,
                column,
            ).Value

            if value is None:
                continue

            if (
                str(value)
                .strip()
                .lower()
                == "результат"
            ):
                return column

        result_column = max_column + 1

        sheet.Cells(
            1,
            result_column,
        ).Value = "Результат"

        return result_column

    # =========================================================
    # SELECT ROW
    # =========================================================

    def select_row(
        self,
        excel_path: str | Path,
        excel_row: int,
    ) -> None:
        if not self.enabled:
            return

        if excel_row < 2:
            return

        with self._lock:
            pythoncom.CoInitialize()

            try:
                app = self._connect()

                workbook = self._find_workbook(
                    app,
                    Path(excel_path),
                )

                workbook.Activate()

                sheet = self._get_first_sheet(
                    workbook
                )

                sheet.Activate()

                cell = sheet.Cells(
                    excel_row,
                    5,
                )

                cell.Select()

                try:
                    app.Goto(
                        cell,
                        True,
                    )
                except Exception:
                    pass

                logger.info(
                    "WPS: cursor moved to E%s",
                    excel_row,
                )

            except Exception as exc:
                logger.warning(
                    "WPS: failed to move cursor "
                    "| row=%s | error=%s",
                    excel_row,
                    exc,
                )

            finally:
                pythoncom.CoUninitialize()

    # =========================================================
    # WRITE RESULT
    # =========================================================

    def write_result(
        self,
        excel_path: str | Path,
        excel_row: int,
        result: str,
    ) -> bool:
        """
        Записывает результат непосредственно
        в уже открытую WPS-книгу.

        Возвращает:
            True  -> записано
            False -> ошибка
        """

        if not self.enabled:
            return False

        if excel_row < 2:
            raise ValueError(
                f"Некорректная строка Excel: {excel_row}"
            )

        with self._lock:
            pythoncom.CoInitialize()

            try:
                app = self._connect()

                workbook = self._find_workbook(
                    app,
                    Path(excel_path),
                )

                workbook.Activate()

                sheet = self._get_first_sheet(
                    workbook
                )

                sheet.Activate()

                result_column = (
                    self._find_result_column(
                        sheet
                    )
                )

                cell = sheet.Cells(
                    excel_row,
                    result_column,
                )

                cell.Value = result

                # Сохраняем открытую книгу.
                workbook.Save()

                # Показываем изменённую строку.
                cell.Select()

                try:
                    app.Goto(
                        cell,
                        True,
                    )
                except Exception:
                    pass

                logger.info(
                    "WPS: result saved | row=%s "
                    "| column=%s | result=%s",
                    excel_row,
                    result_column,
                    result,
                )

                return True

            except Exception:
                logger.exception(
                    "WPS: failed to write result "
                    "| row=%s | result=%s",
                    excel_row,
                    result,
                )

                return False

            finally:
                pythoncom.CoUninitialize()