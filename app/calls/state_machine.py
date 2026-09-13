from __future__ import annotations

import logging
import threading
from datetime import datetime, timezone
from typing import Callable

from app.microsip.controller import MicroSIPController
from app.storage.database import Database

from .classifier import classify_call_duration
from .models import CallResult, CallState


logger = logging.getLogger("ZIK.StateMachine")


# После /hangupall ждём подтверждение END.
# Это не время звонка, а защитный timeout,
# если MicroSIP не отправит END.
ENDING_TIMEOUT_SECONDS = 3.0


class CallStateMachine:
    def __init__(
        self,
        database: Database,
        microsip: MicroSIPController,
        ring_timeout_seconds: int = 30,
    ) -> None:
        self.database = database
        self.microsip = microsip
        self.ring_timeout_seconds = ring_timeout_seconds

        self._watchdogs: dict[int, threading.Timer] = {}
        self._ending_timers: dict[int, threading.Timer] = {}

        self._watchdog_lock = threading.Lock()

        # UI / Controller callback.
        self._state_callback: (
            Callable[[CallState], None] | None
        ) = None

        # Dialer callback.
        self._call_finished_callback: (
            Callable[[str, CallResult], None] | None
        ) = None

    # =========================================================
    # CALLBACKS
    # =========================================================

    def set_state_callback(
        self,
        callback: Callable[[CallState], None] | None,
    ) -> None:
        """
        Регистрирует callback изменения состояния звонка.

        Например:

            OUTGOING
            IN_CALL
            ENDING
            ENDED
        """
        self._state_callback = callback

    def _notify_state(
        self,
        state: CallState,
    ) -> None:
        callback = self._state_callback

        if callback is None:
            return

        try:
            callback(state)
        except Exception:
            logger.exception(
                "STATE_MACHINE: state callback failed"
            )

    def set_call_finished_callback(
        self,
        callback: Callable[[str, CallResult], None] | None,
    ) -> None:
        """
        Регистрирует callback финального результата звонка.

        Вызывается только после фактического завершения
        обработки звонка.
        """
        self._call_finished_callback = callback

    def _notify_call_finished(
        self,
        phone: str,
        result: CallResult,
    ) -> None:
        callback = self._call_finished_callback

        if callback is None:
            return

        try:
            callback(
                phone,
                result,
            )
        except Exception:
            logger.exception(
                "Failed to notify Dialer | "
                "phone=%s result=%s",
                phone,
                result.value,
            )

    # =========================================================
    # TIME
    # =========================================================

    @staticmethod
    def _now() -> datetime:
        return datetime.now(timezone.utc)

    # =========================================================
    # OUTGOING
    # =========================================================

    def outgoing(
        self,
        phone: str,
    ) -> int:
        active = self.database.get_active_call()

        if active is not None:
            logger.warning(
                "OUTGOING ignored: active call exists | "
                "active_call_id=%s active_phone=%s "
                "active_state=%s received_phone=%s",
                active["id"],
                active["phone"],
                active["state"],
                phone,
            )

            # Не создаём второй call.
            return active["id"]

        phone = phone.strip()

        if not phone:
            raise ValueError(
                "Phone number is empty"
            )

        now = self._now()

        call_id = self.database.create_call(
            phone=phone,
            state=CallState.OUTGOING.value,
            result=CallResult.UNKNOWN.value,
            outgoing_at=now,
        )

        logger.info(
            "STATE: IDLE -> OUTGOING | "
            "call_id=%s phone=%s",
            call_id,
            phone,
        )

        # Уведомляем Controller / UI.
        self._notify_state(
            CallState.OUTGOING
        )

        self._start_watchdog(call_id)

        return call_id

    # =========================================================
    # ANSWER / START
    # =========================================================

    def start(
        self,
        phone: str,
    ) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "START ignored: no active call | phone=%s",
                phone,
            )
            return

        if call["phone"] != phone:
            logger.warning(
                "START ignored: phone mismatch | "
                "active=%s received=%s",
                call["phone"],
                phone,
            )
            return

        if call["state"] != CallState.OUTGOING.value:
            logger.warning(
                "START ignored: invalid state | "
                "call_id=%s state=%s",
                call["id"],
                call["state"],
            )
            return

        self._cancel_watchdog(
            call["id"]
        )

        now = self._now()

        self.database.update_call(
            call["id"],
            state=CallState.IN_CALL.value,
            started_at=now,
        )

        logger.info(
            "STATE: OUTGOING -> IN_CALL | "
            "call_id=%s phone=%s",
            call["id"],
            phone,
        )

        # Уведомляем Controller / UI.
        self._notify_state(
            CallState.IN_CALL
        )

    # =========================================================
    # BUSY
    # =========================================================

    def busy(
        self,
        phone: str,
    ) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "BUSY ignored: no active call | phone=%s",
                phone,
            )
            return

        if call["phone"] != phone:
            logger.warning(
                "BUSY ignored: phone mismatch | "
                "active=%s received=%s",
                call["phone"],
                phone,
            )
            return

        if call["state"] not in (
            CallState.OUTGOING.value,
            CallState.IN_CALL.value,
        ):
            logger.warning(
                "BUSY ignored: invalid state | "
                "call_id=%s state=%s",
                call["id"],
                call["state"],
            )
            return

        previous_state = call["state"]

        self._cancel_watchdog(
            call["id"]
        )

        self._cancel_ending_timer(
            call["id"]
        )

        now = self._now()

        self.database.update_call(
            call["id"],
            state=CallState.ENDED.value,
            result=CallResult.BUSY.value,
            ended_at=now,
        )

        logger.info(
            "STATE: %s -> ENDED | "
            "call_id=%s phone=%s result=BUSY",
            previous_state,
            call["id"],
            phone,
        )

        # Сначала сообщаем о переходе состояния.
        self._notify_state(
            CallState.ENDED
        )

        # Затем передаём результат Dialer.
        self._notify_call_finished(
            phone,
            CallResult.BUSY,
        )

    # =========================================================
    # END
    # =========================================================

    def end(
        self,
        phone: str,
    ) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.info(
                "END ignored: no active call | phone=%s",
                phone,
            )
            return

        if call["phone"] != phone:
            logger.warning(
                "END ignored: phone mismatch | "
                "active=%s received=%s",
                call["phone"],
                phone,
            )
            return

        self._cancel_watchdog(
            call["id"]
        )

        self._cancel_ending_timer(
            call["id"]
        )

        if call["state"] == CallState.ENDED.value:
            logger.info(
                "END ignored: already finalized | "
                "call_id=%s phone=%s",
                call["id"],
                phone,
            )
            return

        now = self._now()

        # -----------------------------------------------------
        # WATCHDOG -> ENDING -> END
        # -----------------------------------------------------

        if call["state"] == CallState.ENDING.value:
            result = CallResult.NO_ANSWER

            self.database.update_call(
                call["id"],
                state=CallState.ENDED.value,
                result=result.value,
                ended_at=now,
            )

            logger.info(
                "STATE: ENDING -> ENDED | "
                "call_id=%s phone=%s result=NO_ANSWER",
                call["id"],
                phone,
            )

            self._notify_state(
                CallState.ENDED
            )

            self._notify_call_finished(
                phone,
                result,
            )

            return

        # -----------------------------------------------------
        # NORMAL ANSWERED CALL
        # -----------------------------------------------------

        if call["manual_result"]:
            result = CallResult(
                call["manual_result"]
            )

        elif call["started_at"]:
            started_at = datetime.fromisoformat(
                call["started_at"]
            )

            if started_at.tzinfo is None:
                started_at = started_at.replace(
                    tzinfo=timezone.utc
                )

            duration = (
                now - started_at
            ).total_seconds()

            logger.info(
                "Call duration: %.3f sec | "
                "call_id=%s phone=%s",
                duration,
                call["id"],
                phone,
            )

            result = classify_call_duration(
                duration
            )

        else:
            result = CallResult.NO_ANSWER

        previous_state = call["state"]

        self.database.update_call(
            call["id"],
            state=CallState.ENDED.value,
            result=result.value,
            ended_at=now,
        )

        logger.info(
            "STATE: %s -> ENDED | "
            "call_id=%s phone=%s result=%s",
            previous_state,
            call["id"],
            phone,
            result.value,
        )

        self._notify_state(
            CallState.ENDED
        )

        self._notify_call_finished(
            phone,
            result,
        )

    # =========================================================
    # KOCHADAN
    # =========================================================

    def kochadan(self) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "KOCHADAN ignored: no active call"
            )
            return

        if call["state"] != CallState.IN_CALL.value:
            logger.warning(
                "KOCHADAN ignored: invalid state | "
                "call_id=%s state=%s",
                call["id"],
                call["state"],
            )
            return

        self.database.update_call(
            call["id"],
            manual_result=CallResult.KOCHADAN.value,
            result=CallResult.KOCHADAN.value,
        )

        logger.info(
            "MANUAL RESULT: KOCHADAN | "
            "call_id=%s phone=%s",
            call["id"],
            call["phone"],
        )

    # =========================================================
    # WATCHDOG
    # =========================================================

    def _start_watchdog(
        self,
        call_id: int,
    ) -> None:
        timer = threading.Timer(
            self.ring_timeout_seconds,
            self._watchdog_expired,
            args=(call_id,),
        )

        timer.daemon = True

        with self._watchdog_lock:
            old_timer = self._watchdogs.pop(
                call_id,
                None,
            )

            if old_timer is not None:
                old_timer.cancel()

            self._watchdogs[call_id] = timer

        timer.start()

        logger.info(
            "WATCHDOG: started | "
            "call_id=%s timeout=%s sec",
            call_id,
            self.ring_timeout_seconds,
        )

    def _cancel_watchdog(
        self,
        call_id: int,
    ) -> None:
        with self._watchdog_lock:
            timer = self._watchdogs.pop(
                call_id,
                None,
            )

        if timer is not None:
            timer.cancel()

            logger.info(
                "WATCHDOG: cancelled | "
                "call_id=%s",
                call_id,
            )

    # =========================================================
    # ENDING TIMER
    # =========================================================

    def _start_ending_timer(
        self,
        call_id: int,
    ) -> None:
        timer = threading.Timer(
            ENDING_TIMEOUT_SECONDS,
            self._ending_timeout,
            args=(call_id,),
        )

        timer.daemon = True

        with self._watchdog_lock:
            old_timer = self._ending_timers.pop(
                call_id,
                None,
            )

            if old_timer is not None:
                old_timer.cancel()

            self._ending_timers[call_id] = timer

        timer.start()

        logger.info(
            "ENDING timer started | "
            "call_id=%s timeout=%.1f sec",
            call_id,
            ENDING_TIMEOUT_SECONDS,
        )

    def _cancel_ending_timer(
        self,
        call_id: int,
    ) -> None:
        with self._watchdog_lock:
            timer = self._ending_timers.pop(
                call_id,
                None,
            )

        if timer is not None:
            timer.cancel()

    def _ending_timeout(
        self,
        call_id: int,
    ) -> None:
        call = self.database.get_call(
            call_id
        )

        if call is None:
            return

        if call["state"] != CallState.ENDING.value:
            return

        now = self._now()

        logger.warning(
            "ENDING timeout: MicroSIP END not received | "
            "call_id=%s phone=%s",
            call_id,
            call["phone"],
        )

        self.database.update_call(
            call_id,
            state=CallState.ENDED.value,
            result=CallResult.NO_ANSWER.value,
            ended_at=now,
        )

        logger.info(
            "STATE: ENDING -> ENDED by fallback | "
            "call_id=%s phone=%s result=NO_ANSWER",
            call_id,
            call["phone"],
        )

        self._notify_state(
            CallState.ENDED
        )

        self._notify_call_finished(
            call["phone"],
            CallResult.NO_ANSWER,
        )

        with self._watchdog_lock:
            self._ending_timers.pop(
                call_id,
                None,
            )

    # =========================================================
    # WATCHDOG EXPIRED
    # =========================================================

    def _watchdog_expired(
        self,
        call_id: int,
    ) -> None:
        logger.info(
            "WATCHDOG: timeout reached | "
            "call_id=%s",
            call_id,
        )

        call = self.database.get_call(
            call_id
        )

        if call is None:
            logger.warning(
                "WATCHDOG: call not found | "
                "call_id=%s",
                call_id,
            )
            return

        if call["state"] != CallState.OUTGOING.value:
            logger.info(
                "WATCHDOG: nothing to do | "
                "call_id=%s state=%s",
                call_id,
                call["state"],
            )
            return

        logger.warning(
            "WATCHDOG: no answer after %s sec | "
            "call_id=%s phone=%s",
            self.ring_timeout_seconds,
            call_id,
            call["phone"],
        )

        # -----------------------------------------------------
        # OUTGOING -> ENDING
        #
        # Пока MicroSIP не подтвердил END,
        # Dialer ещё НЕ получает результат.
        # -----------------------------------------------------

        self.database.update_call(
            call_id,
            state=CallState.ENDING.value,
            result=CallResult.NO_ANSWER.value,
        )

        logger.info(
            "STATE: OUTGOING -> ENDING | "
            "call_id=%s phone=%s result=NO_ANSWER",
            call_id,
            call["phone"],
        )

        self._notify_state(
            CallState.ENDING
        )

        try:
            self.microsip.hangup_all()

            logger.info(
                "WATCHDOG: MicroSIP /hangupall sent | "
                "call_id=%s",
                call_id,
            )

        except Exception:
            logger.exception(
                "WATCHDOG: failed to hang up MicroSIP | "
                "call_id=%s",
                call_id,
            )

        # Даём MicroSIP время отправить END.
        self._start_ending_timer(
            call_id
        )

        with self._watchdog_lock:
            self._watchdogs.pop(
                call_id,
                None,
            )

    # =========================================================
    # SHUTDOWN
    # =========================================================

    def shutdown(self) -> None:
        logger.info(
            "STATE: shutting down"
        )

        # -----------------------------------------------------
        # Cancel watchdogs / ending timers
        # -----------------------------------------------------

        with self._watchdog_lock:
            watchdogs = list(
                self._watchdogs.values()
            )

            ending_timers = list(
                self._ending_timers.values()
            )

            self._watchdogs.clear()
            self._ending_timers.clear()

        for timer in watchdogs:
            timer.cancel()

        for timer in ending_timers:
            timer.cancel()

        # -----------------------------------------------------
        # Finalize active call
        # -----------------------------------------------------

        call = self.database.get_active_call()

        if call is None:
            logger.info(
                "STATE: shutdown complete | "
                "no active call"
            )
            return

        logger.warning(
            "STATE: finalizing active call on shutdown | "
            "call_id=%s phone=%s state=%s",
            call["id"],
            call["phone"],
            call["state"],
        )

        now = self._now()

        self.database.update_call(
            call["id"],
            state=CallState.ENDED.value,
            result=CallResult.NO_ANSWER.value,
            ended_at=now,
        )

        self._notify_state(
            CallState.ENDED
        )

        logger.info(
            "STATE: active call finalized | "
            "call_id=%s result=NO_ANSWER",
            call["id"],
        )