from __future__ import annotations

import logging
import threading

from app.calls.models import CallResult
from app.excel.writer import ExcelWriter
from app.microsip.controller import MicroSIPController

from .queue import CallQueue, MAX_ATTEMPTS


logger = logging.getLogger("ZIK.Dialer")


class Dialer:

    def __init__(
        self,
        queue: CallQueue,
        microsip: MicroSIPController,
        excel_writer: ExcelWriter | None = None,
        retry_delay_seconds: float = 0.5,
    ) -> None:
        self.queue = queue
        self.microsip = microsip
        self.excel_writer = excel_writer
        self.retry_delay_seconds = retry_delay_seconds

        self._lock = threading.RLock()

        self._running = False
        self._paused = False

    # =========================================================
    # START
    # =========================================================

    def start(self) -> None:
        with self._lock:
            if self._running:
                return

            self._running = True
            self._paused = False

        logger.info(
            "DIALER: started | queue_size=%s",
            self.queue.size(),
        )

        self._dial_next()

    # =========================================================
    # STOP
    # =========================================================

    def stop(self) -> None:
        with self._lock:
            if not self._running:
                return

            self._running = False
            self._paused = False

        logger.info("DIALER: stopping")

        try:
            self.microsip.hangup_all()

            logger.info(
                "DIALER: MicroSIP calls terminated"
            )

        except Exception:
            logger.exception(
                "DIALER: failed to terminate MicroSIP"
            )

        logger.info("DIALER: stopped")

    # =========================================================
    # PAUSE
    # =========================================================

    def pause(self) -> None:
        with self._lock:
            if not self._running:
                return

            self._paused = True

        logger.info(
            "DIALER: paused"
        )

    # =========================================================
    # RESUME
    # =========================================================

    def resume(self) -> None:
        with self._lock:
            if not self._running:
                return

            self._paused = False

        logger.info(
            "DIALER: resumed"
        )

        self._dial_next()

    # =========================================================
    # STATE
    # =========================================================

    @property
    def is_running(self) -> bool:
        with self._lock:
            return self._running

    @property
    def is_paused(self) -> bool:
        with self._lock:
            return self._paused

    # =========================================================
    # DIAL
    # =========================================================

    def _dial_next(self) -> None:
        with self._lock:
            if not self._running:
                return

            if self._paused:
                logger.info(
                    "DIALER: next call blocked because "
                    "dialer is paused"
                )
                return

            item = self.queue.current

            if item is None:
                item = self.queue.start_next()

            if item is None:
                self._running = False

                logger.info(
                    "DIALER: queue is empty"
                )

                return

            attempt = self.queue.mark_attempt()

        logger.info(
            "DIALER: dialing | phone=%s "
            "excel_row=%s attempt=%s/%s",
            item.phone,
            item.excel_row,
            attempt,
            MAX_ATTEMPTS,
        )

        try:
            self.microsip.call(
                item.phone
            )

        except Exception:
            logger.exception(
                "DIALER: failed to start call "
                "| phone=%s",
                item.phone,
            )

            self.queue.complete_current()

            self._schedule_next()

    # =========================================================
    # RESULT
    # =========================================================

    def handle_result(
        self,
        phone: str,
        result: CallResult,
    ) -> None:

        with self._lock:
            if not self._running:
                logger.info(
                    "DIALER: result ignored "
                    "because dialer is stopped "
                    "| phone=%s | result=%s",
                    phone,
                    result.value,
                )
                return

            current = self.queue.current

            if current is None:
                logger.warning(
                    "DIALER: result without "
                    "current queue item "
                    "| phone=%s",
                    phone,
                )
                return

            if current.phone != phone:
                logger.warning(
                    "DIALER: phone mismatch "
                    "| queue=%s | result=%s",
                    current.phone,
                    phone,
                )
                return

            logger.info(
                "DIALER: result | phone=%s "
                "excel_row=%s attempt=%s/%s "
                "result=%s",
                phone,
                current.excel_row,
                current.attempts,
                MAX_ATTEMPTS,
                result.value,
            )

            # -------------------------------------------------
            # FAILED CALL
            # -------------------------------------------------

            if result in (
                CallResult.NO_ANSWER,
                CallResult.BUSY,
            ):

                if self.queue.retry_current():

                    logger.info(
                        "DIALER: %s -> retry | "
                        "phone=%s | next attempt=%s/%s",
                        result.value,
                        phone,
                        current.attempts + 1,
                        MAX_ATTEMPTS,
                    )

                    self._schedule_next()

                    return

                # Максимум достигнут.
                logger.info(
                    "DIALER: maximum attempts reached "
                    "| phone=%s | final=%s",
                    phone,
                    result.value,
                )

                self._write_excel_result(
                    current.excel_row,
                    result,
                )

                self.queue.complete_current()

                self._schedule_next()

                return

            # -------------------------------------------------
            # SUCCESSFUL RESULT
            # -------------------------------------------------

            logger.info(
                "DIALER: phone completed | "
                "phone=%s | row=%s | result=%s",
                phone,
                current.excel_row,
                result.value,
            )

            self._write_excel_result(
                current.excel_row,
                result,
            )

            self.queue.complete_current()

            self._schedule_next()

    # =========================================================
    # EXCEL
    # =========================================================

    def _write_excel_result(
        self,
        excel_row: int | None,
        result: CallResult,
    ) -> None:

        if excel_row is None:
            logger.warning(
                "DIALER: excel_row is None "
                "| result=%s",
                result.value,
            )
            return

        if self.excel_writer is None:
            logger.warning(
                "DIALER: ExcelWriter is not configured"
            )
            return

        try:
            self.excel_writer.write_result(
                excel_row,
                result.value,
            )

            logger.info(
                "DIALER: Excel result saved "
                "| row=%s | result=%s",
                excel_row,
                result.value,
            )

        except Exception:
            logger.exception(
                "DIALER: failed to save Excel "
                "| row=%s | result=%s",
                excel_row,
                result.value,
            )

    # =========================================================
    # NEXT
    # =========================================================

    def _schedule_next(self) -> None:
        timer = threading.Timer(
            self.retry_delay_seconds,
            self._dial_next,
        )

        timer.daemon = True
        timer.start()