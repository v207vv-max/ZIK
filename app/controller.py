from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Callable
from app.excel.wps_sync import WPSSync
from app.calls.models import CallResult, CallState
from app.calls.state_machine import CallStateMachine
from app.dialer.dialer import Dialer
from app.dialer.excel_source import build_call_items
from app.dialer.queue import CallQueue
from app.excel.filter import ResultFilter
from app.excel.reader import ExcelReader
from app.excel.writer import ExcelWriter
from app.microsip.controller import MicroSIPController


logger = logging.getLogger("ZIK.Controller")


class ZIKController:
    """
    Центральный контроллер ZIK.

    UI не работает напрямую с:
        - MicroSIP
        - StateMachine
        - CallQueue
        - ExcelWriter

    UI обращается только к этому классу.
    """

    def __init__(
        self,
        microsip: MicroSIPController,
        state_machine: CallStateMachine,
    ) -> None:
        self.wps = WPSSync()
        self.writer = writer
        self.microsip = microsip
        self.state_machine = state_machine

        self.excel_path: Path | None = None

        self.reader: ExcelReader | None = None
        self.writer: ExcelWriter | None = None
        self.result_filter: ResultFilter | None = None

        self.queue = CallQueue()

        self.dialer = Dialer(
            queue=self.queue,
            microsip=self.microsip,
            excel_writer=None,
        )

        self.state_machine.set_call_finished_callback(
            self._handle_call_result
        )

        self._lock = threading.RLock()

        self._running = False
        self._paused = False

        self._current_phone = ""
        self._current_row: int | None = None
        self._current_attempt = 0

        self._state = CallState.IDLE
        self._last_result: CallResult | None = None

        self._status_callback: Callable[[dict], None] | None = None
        self._log_callback: Callable[[str], None] | None = None

    # ============================================================
    # CALLBACKS
    # ============================================================

    def set_status_callback(
        self,
        callback: Callable[[dict], None] | None,
    ) -> None:
        self._status_callback = callback

    def set_log_callback(
        self,
        callback: Callable[[str], None] | None,
    ) -> None:
        self._log_callback = callback

    def _emit_status(self) -> None:
        callback = self._status_callback

        if callback is None:
            return

        with self._lock:
            item = self.queue.current

            status = {
                "running": self._running,
                "paused": self._paused,
                "state": self._state.value,
                "phone": (
                    item.phone
                    if item is not None
                    else self._current_phone
                ),
                "excel_row": (
                    item.excel_row
                    if item is not None
                    else self._current_row
                ),
                "attempt": (
                    item.attempts
                    if item is not None
                    else self._current_attempt
                ),
                "max_attempts": 3,
                "result": (
                    self._last_result.value
                    if self._last_result is not None
                    else ""
                ),
                "queue_size": self.queue.size(),
            }

        try:
            callback(status)
        except Exception:
            logger.exception(
                "CONTROLLER: status callback failed"
            )

    def _emit_log(self, message: str) -> None:
        logger.info(message)

        callback = self._log_callback

        if callback is None:
            return

        try:
            callback(message)
        except Exception:
            logger.exception(
                "CONTROLLER: log callback failed"
            )

    # ============================================================
    # EXCEL
    # ============================================================

    def load_excel(self, path: str | Path) -> list:
        """
        Загружает Excel и подготавливает ResultFilter.
        """

        excel_path = Path(path)

        if not excel_path.exists():
            raise FileNotFoundError(
                f"Excel-файл не найден: {excel_path}"
            )

        reader = ExcelReader(excel_path)
        rows = reader.load()

        if not rows:
            raise RuntimeError(
                "Excel не содержит строк."
            )

        writer = ExcelWriter(excel_path)
        result_filter = ResultFilter(rows)

        with self._lock:
            self.excel_path = excel_path
            self.reader = reader
            self.writer = writer
            self.result_filter = result_filter

            # Подключаем Writer к уже существующему Dialer.
            self.dialer.excel_writer = writer

        self._emit_log(
            f"Excel загружен: {excel_path}"
        )

        self._emit_log(
            f"Строк загружено: {len(rows)}"
        )

        self._emit_status()

        return rows

    def get_excel_rows(self) -> list:
        with self._lock:
            if self.reader is None:
                return []

            return self.reader.load()

    def set_result_filter(
        self,
        selected: set[str] | None,
    ) -> None:
        with self._lock:
            if self.result_filter is None:
                raise RuntimeError(
                    "Сначала загрузите Excel."
                )

            self.result_filter.set_selected(selected)

    def clear_result_filter(self) -> None:
        with self._lock:
            if self.result_filter is not None:
                self.result_filter.clear()

    # ============================================================
    # QUEUE
    # ============================================================

    def prepare_queue(
        self,
        start_row: int | None = None,
    ) -> int:
        """
        Формирует очередь звонков.

        Если start_row указан:
            обработка начинается с этой строки Excel.
        """

        with self._lock:
            if self.reader is None:
                raise RuntimeError(
                    "Excel не загружен."
                )

            rows = self.reader.load()

            if self.result_filter is None:
                filtered_rows = rows
            else:
                filtered_rows = self.result_filter.apply(rows)

            if start_row is not None:
                filtered_rows = [
                    row
                    for row in filtered_rows
                    if row.excel_row >= start_row
                ]

            items = build_call_items(filtered_rows)

            self.queue.clear()
            self.queue.add_many(items)

            count = self.queue.size()

        self._emit_log(
            f"Очередь подготовлена: {count} номеров"
        )

        if start_row is not None:
            self._emit_log(
                f"Начальная строка Excel: {start_row}"
            )

        self._emit_status()

        return count

    # ============================================================
    # START
    # ============================================================

    def start(
        self,
        start_row: int | None = None,
    ) -> None:
        with self._lock:
            if self._running:
                self._emit_log(
                    "START: автодозвон уже запущен."
                )
                return

        self.prepare_queue(start_row)

        if not self.queue.has_items():
            raise RuntimeError(
                "После фильтрации очередь пуста."
            )

        with self._lock:
            self._running = True
            self._paused = False
            self._last_result = None

        self._emit_log("START: автодозвон запущен.")

        self.dialer.start()

        self._emit_status()

    # ============================================================
    # PAUSE
    # ============================================================

    def pause(self) -> None:
        with self._lock:
            if not self._running:
                return

            if self._paused:
                return

            self._paused = True

        self.dialer.pause()

        self._emit_log(
            "PAUSE: автодозвон поставлен на паузу."
        )

        self._emit_status()

    # ============================================================
    # RESUME
    # ============================================================

    def resume(self) -> None:
        with self._lock:
            if not self._running:
                return

            if not self._paused:
                return

            self._paused = False

        self.dialer.resume()

        self._emit_log(
            "RESUME: автодозвон продолжен."
        )

        self._emit_status()

    # ============================================================
    # STOP
    # ============================================================

    def stop(self) -> None:
        with self._lock:
            was_running = self._running

            self._running = False
            self._paused = False

        self.dialer.stop()

        if was_running:
            try:
                self.microsip.hangup_all()
            except Exception:
                logger.exception(
                    "CONTROLLER: failed to hang up MicroSIP"
                )

        self._emit_log(
            "STOP: автодозвон остановлен."
        )

        self._emit_status()

    # ============================================================
    # RESULT
    # ============================================================

    def _handle_call_result(
        self,
        phone: str,
        result: CallResult,
    ) -> None:

        with self._lock:
            self._last_result = result

            item = self.queue.current

            if item is not None:
                self._current_phone = item.phone
                self._current_row = item.excel_row
                self._current_attempt = item.attempts

            excel_path = self.excel_path
            excel_row = (
                item.excel_row
                if item is not None
                else self._current_row
            )

        self._emit_log(
            f"Результат звонка: "
            f"{phone} -> {result.value}"
        )

        # =========================================================
        # WPS
        # =========================================================

        if (
            excel_path is not None
            and excel_row is not None
        ):
            success = self.wps.write_result(
                excel_path,
                excel_row,
                result.value,
            )

            if success:
                self._emit_log(
                    f"WPS: результат записан | "
                    f"row={excel_row} | "
                    f"result={result.value}"
                )

            else:
                self._emit_log(
                    f"WPS: НЕ удалось записать результат | "
                    f"row={excel_row}"
                )

    # =========================================================
    # DIALER
    # =========================================================

    self.dialer.handle_result(
        phone,
        result,
    )

    # =========================================================
    # UI
    # =========================================================

    self._emit_status()

    # Даём Dialer завершить retry/current
    # и получить следующий QueueItem.
    threading.Timer(
        0.05,
        self._emit_status,
    ).start()
    # ============================================================
    # STATE
    # ============================================================

    def update_state(
        self,
        state: CallState,
    ) -> None:
        with self._lock:
            self._state = state

        self._emit_status()

    # ============================================================
    # CURRENT ITEM
    # ============================================================

    def get_current(self) -> dict:
        with self._lock:
            item = self.queue.current

            if item is None:
                return {
                    "phone": self._current_phone,
                    "excel_row": self._current_row,
                    "attempt": self._current_attempt,
                    "state": self._state.value,
                    "result": (
                        self._last_result.value
                        if self._last_result
                        else ""
                    ),
                }

            return {
                "phone": item.phone,
                "excel_row": item.excel_row,
                "attempt": item.attempts,
                "state": self._state.value,
                "result": (
                    self._last_result.value
                    if self._last_result
                    else ""
                ),
            }

    # ============================================================
    # PROPERTIES
    # ============================================================

    @property
    def is_running(self) -> bool:
        with self._lock:
            return self._running

    @property
    def is_paused(self) -> bool:
        with self._lock:
            return self._paused

    @property
    def current_phone(self) -> str:
        return self.get_current()["phone"]

    @property
    def current_row(self) -> int | None:
        return self.get_current()["excel_row"]

    @property
    def current_attempt(self) -> int:
        return int(self.get_current()["attempt"])

    @property
    def current_state(self) -> str:
        return self.get_current()["state"]

    @property
    def queue_size(self) -> int:
        return self.queue.size()