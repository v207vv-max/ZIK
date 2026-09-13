from __future__ import annotations

import logging
import threading
from datetime import datetime, timezone

from app.microsip.controller import MicroSIPController
from app.storage.database import Database

from .classifier import classify_call_duration
from .models import CallResult, CallState

logger = logging.getLogger("ZIK.StateMachine")


class CallStateMachine:
    def __init__(
        self,
        database: Database,
        microsip: MicroSIPController,
        ring_timeout_seconds: int = 25,
    ) -> None:
        self.database = database
        self.microsip = microsip
        self.ring_timeout_seconds = ring_timeout_seconds

        self._watchdogs: dict[int, threading.Timer] = {}
        self._watchdog_lock = threading.Lock()

    @staticmethod
    def _now() -> datetime:
        return datetime.now(timezone.utc)

    def outgoing(self, phone: str) -> int:
        active = self.database.get_active_call()

        if active is not None:
            logger.warning(
                "Cannot start outgoing call: active call exists | "
                "call_id=%s phone=%s",
                active["id"],
                active["phone"],
            )
            raise RuntimeError("An active call already exists")

        now = self._now()

        call_id = self.database.create_call(
            phone=phone,
            state=CallState.OUTGOING.value,
            result=CallResult.UNKNOWN.value,
            outgoing_at=now.isoformat(),
        )

        logger.info(
            "STATE: IDLE -> OUTGOING | call_id=%s phone=%s",
            call_id,
            phone,
        )

        self._start_watchdog(call_id)

        return call_id

    def start(self, phone: str) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "START received but there is no active call | phone=%s",
                phone,
            )
            return

        if call["phone"] != phone:
            logger.warning(
                "START phone mismatch | active=%s received=%s",
                call["phone"],
                phone,
            )
            return

        if call["state"] != CallState.OUTGOING.value:
            logger.warning(
                "START ignored | call_id=%s state=%s",
                call["id"],
                call["state"],
            )
            return

        self._cancel_watchdog(call["id"])

        now = self._now()

        self.database.update_call(
            call["id"],
            state=CallState.IN_CALL.value,
            started_at=now.isoformat(),
        )

        logger.info(
            "STATE: OUTGOING -> IN_CALL | call_id=%s phone=%s",
            call["id"],
            phone,
        )

    def busy(self, phone: str) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "BUSY received but there is no active call | phone=%s",
                phone,
            )
            return

        if call["phone"] != phone:
            logger.warning(
                "BUSY phone mismatch | active=%s received=%s",
                call["phone"],
                phone,
            )
            return

        self._cancel_watchdog(call["id"])

        now = self._now()

        self.database.update_call(
            call["id"],
            state=CallState.BUSY.value,
            result=CallResult.BUSY.value,
            ended_at=now.isoformat(),
        )

        logger.info(
            "STATE: %s -> BUSY | call_id=%s phone=%s",
            call["state"],
            call["id"],
            phone,
        )

    def end(self, phone: str) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "END received but there is no active call | phone=%s",
                phone,
            )
            return

        if call["phone"] != phone:
            logger.warning(
                "END phone mismatch | active=%s received=%s",
                call["phone"],
                phone,
            )
            return

        self._cancel_watchdog(call["id"])

        now = self._now()

        # Manual result always has priority.
        if call["manual_result"]:
            result = CallResult(call["manual_result"])

        # Answered call.
        elif call["started_at"]:
            started_at = datetime.fromisoformat(call["started_at"])
            duration = (now - started_at).total_seconds()

            logger.info(
                "Call duration: %.3f sec | call_id=%s",
                duration,
                call["id"],
            )

            result = classify_call_duration(duration)

        # Ended before answer.
        else:
            result = CallResult.NO_ANSWER

        self.database.update_call(
            call["id"],
            state=CallState.ENDED.value,
            result=result.value,
            ended_at=now.isoformat(),
        )

        logger.info(
            "STATE: %s -> ENDED | call_id=%s phone=%s result=%s",
            call["state"],
            call["id"],
            phone,
            result.value,
        )

    def kochadan(self) -> None:
        call = self.database.get_active_call()

        if call is None:
            logger.warning(
                "KOCHADAN received but there is no active call"
            )
            return

        if call["state"] != CallState.IN_CALL.value:
            logger.warning(
                "KOCHADAN ignored | call_id=%s state=%s",
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
            "MANUAL RESULT: KOCHADAN | call_id=%s phone=%s",
            call["id"],
            call["phone"],
        )

    # ---------------------------------------------------------
    # WATCHDOG
    # ---------------------------------------------------------

    def _start_watchdog(self, call_id: int) -> None:
        timer = threading.Timer(
            self.ring_timeout_seconds,
            self._watchdog_expired,
            args=(call_id,),
        )

        timer.daemon = True

        with self._watchdog_lock:
            old_timer = self._watchdogs.pop(call_id, None)

            if old_timer is not None:
                old_timer.cancel()

            self._watchdogs[call_id] = timer

        timer.start()

        logger.info(
            "WATCHDOG: started | call_id=%s timeout=%s sec",
            call_id,
            self.ring_timeout_seconds,
        )

    def _cancel_watchdog(self, call_id: int) -> None:
        with self._watchdog_lock:
            timer = self._watchdogs.pop(call_id, None)

        if timer is not None:
            timer.cancel()

            logger.info(
                "WATCHDOG: cancelled | call_id=%s",
                call_id,
            )

    def _watchdog_expired(self, call_id: int) -> None:
        logger.info(
            "WATCHDOG: timeout reached | call_id=%s",
            call_id,
        )

        # IMPORTANT:
        # Re-read the database. We must not trust the state
        # captured when the timer was created.
        call = self.database.get_call(call_id)

        if call is None:
            logger.warning(
                "WATCHDOG: call disappeared | call_id=%s",
                call_id,
            )
            return

        # Race protection:
        # if START arrived at approximately the same time,
        # the database will already say IN_CALL.
        if call["state"] != CallState.OUTGOING.value:
            logger.info(
                "WATCHDOG: nothing to do | call_id=%s state=%s",
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

        # Hang up MicroSIP first.
        try:
            self.microsip.hangup_all()

            logger.info(
                "WATCHDOG: MicroSIP /hangupall sent | call_id=%s",
                call_id,
            )

        except Exception:
            logger.exception(
                "WATCHDOG: failed to hang up MicroSIP | call_id=%s",
                call_id,
            )

        # Mark this exact call as NO_ANSWER.
        now = self._now()

        self.database.update_call(
            call_id,
            state=CallState.ENDED.value,
            result=CallResult.NO_ANSWER.value,
            ended_at=now.isoformat(),
        )

        logger.info(
            "STATE: OUTGOING -> ENDED | "
            "call_id=%s phone=%s result=NO_ANSWER",
            call_id,
            call["phone"],
        )

        with self._watchdog_lock:
            self._watchdogs.pop(call_id, None)