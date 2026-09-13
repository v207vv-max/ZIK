from __future__ import annotations

import json
import logging
import threading
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from queue import Queue, Empty

from .calls.state_machine import CallStateMachine


logger = logging.getLogger("ZIK.IPC")


# ============================================================
# EVENT
# ============================================================


@dataclass(frozen=True)
class IPCEvent:
    event: str
    phone: str
    timestamp: object = None


# ============================================================
# EVENT WORKER
# ============================================================


class IPCEventWorker:
    """
    Последовательно обрабатывает события MicroSIP.

    HTTP-запрос не ждёт выполнения StateMachine.
    Событие сначала попадает в очередь, после чего
    HTTP сразу получает 200 OK.

    Один worker гарантирует порядок:

        outgoing
        start
        end

    """

    def __init__(
        self,
        state_machine: CallStateMachine,
    ) -> None:
        self.state_machine = state_machine

        self._queue: Queue[
            IPCEvent | None
        ] = Queue()

        self._thread = threading.Thread(
            target=self._run,
            name="ZIK-IPC-Worker",
            daemon=True,
        )

        self._running = False

    # --------------------------------------------------------
    # START
    # --------------------------------------------------------

    def start(self) -> None:
        if self._running:
            return

        self._running = True

        self._thread.start()

        logger.info(
            "IPC worker started"
        )

    # --------------------------------------------------------
    # STOP
    # --------------------------------------------------------

    def stop(self) -> None:
        if not self._running:
            return

        self._running = False

        self._queue.put(None)

        if (
            threading.current_thread()
            is not self._thread
        ):
            self._thread.join(
                timeout=2.0
            )

        logger.info(
            "IPC worker stopped"
        )

    # --------------------------------------------------------
    # QUEUE
    # --------------------------------------------------------

    def put(
        self,
        event: IPCEvent,
    ) -> None:
        if not self._running:
            raise RuntimeError(
                "IPC worker is not running"
            )

        self._queue.put(event)

        logger.info(
            "EVENT queued | "
            "event=%s phone=%s queue=%s",
            event.event,
            event.phone,
            self._queue.qsize(),
        )

    # --------------------------------------------------------
    # WORKER LOOP
    # --------------------------------------------------------

    def _run(self) -> None:
        while True:
            try:
                item = self._queue.get()

            except Exception:
                logger.exception(
                    "IPC worker queue error"
                )
                continue

            try:
                if item is None:
                    return

                self._process(item)

            except Exception:
                logger.exception(
                    "IPC worker event error | "
                    "event=%s phone=%s",
                    getattr(item, "event", None),
                    getattr(item, "phone", None),
                )

            finally:
                self._queue.task_done()

    # --------------------------------------------------------
    # PROCESS
    # --------------------------------------------------------

    def _process(
        self,
        event: IPCEvent,
    ) -> None:
        logger.info(
            "EVENT processing | "
            "event=%s phone=%s timestamp=%s",
            event.event,
            event.phone,
            event.timestamp,
        )

        if event.event == "outgoing":
            self.state_machine.outgoing(
                event.phone
            )
            return

        if event.event == "start":
            self.state_machine.start(
                event.phone
            )
            return

        if event.event == "busy":
            self.state_machine.busy(
                event.phone
            )
            return

        if event.event == "end":
            self.state_machine.end(
                event.phone
            )
            return

        if event.event == "kochadan":
            self.state_machine.kochadan(
                event.phone
            )
            return

        logger.warning(
            "EVENT ignored: unknown event=%s phone=%s",
            event.event,
            event.phone,
        )


# ============================================================
# HTTP HANDLER
# ============================================================


class IPCHandler(BaseHTTPRequestHandler):
    """
    Local HTTP IPC endpoint.

    MicroSIP event_receiver.py sends events here:

        POST /event

    Event processing is asynchronous.

    The request is accepted and answered immediately,
    while StateMachine processes events sequentially
    in IPCEventWorker.
    """

    event_worker: IPCEventWorker | None = None

    def log_message(
        self,
        format: str,
        *args,
    ) -> None:
        logger.info(
            "HTTP | " + format,
            *args,
        )

    # --------------------------------------------------------
    # POST
    # --------------------------------------------------------

    def do_POST(self) -> None:
        if self.path != "/event":
            self._send_json(
                404,
                {
                    "ok": False,
                    "error": "Not found",
                },
            )
            return

        # ----------------------------------------------------
        # Content-Length
        # ----------------------------------------------------

        try:
            content_length = int(
                self.headers.get(
                    "Content-Length",
                    "0",
                )
            )

        except ValueError:
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "Invalid Content-Length",
                },
            )
            return

        if content_length <= 0:
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "Empty request body",
                },
            )
            return

        # ----------------------------------------------------
        # Body
        # ----------------------------------------------------

        try:
            raw_body = self.rfile.read(
                content_length
            )

        except Exception:
            logger.exception(
                "HTTP: failed to read request body"
            )

            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "Failed to read request body",
                },
            )
            return

        # ----------------------------------------------------
        # JSON
        # ----------------------------------------------------

        try:
            data = json.loads(
                raw_body.decode("utf-8")
            )

        except (
            UnicodeDecodeError,
            json.JSONDecodeError,
        ):
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "Invalid JSON",
                },
            )
            return

        if not isinstance(data, dict):
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "JSON body must be an object",
                },
            )
            return

        # ----------------------------------------------------
        # EVENT
        # ----------------------------------------------------

        event = str(
            data.get(
                "event",
                "",
            )
        ).strip().lower()

        phone = str(
            data.get(
                "phone",
                "",
            )
        ).strip()

        timestamp = data.get(
            "timestamp"
        )

        logger.info(
            "EVENT received | "
            "event=%s phone=%s timestamp=%s",
            event,
            phone,
            timestamp,
        )

        if not event:
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "Missing event",
                },
            )
            return

        # ----------------------------------------------------
        # WORKER
        # ----------------------------------------------------

        worker = self.event_worker

        if worker is None:
            self._send_json(
                503,
                {
                    "ok": False,
                    "error": "IPC worker is not initialized",
                },
            )
            return

        if not worker._running:
            self._send_json(
                503,
                {
                    "ok": False,
                    "error": "IPC worker is stopped",
                },
            )
            return

        # ----------------------------------------------------
        # QUEUE EVENT
        # ----------------------------------------------------

        try:
            worker.put(
                IPCEvent(
                    event=event,
                    phone=phone,
                    timestamp=timestamp,
                )
            )

        except Exception:
            logger.exception(
                "EVENT queue error | "
                "event=%s phone=%s",
                event,
                phone,
            )

            self._send_json(
                500,
                {
                    "ok": False,
                    "error": "Failed to queue event",
                },
            )
            return

        # ----------------------------------------------------
        # IMPORTANT:
        #
        # Respond immediately.
        #
        # We do NOT wait for StateMachine.
        # ----------------------------------------------------

        self._send_json(
            200,
            {
                "ok": True,
                "queued": True,
                "event": event,
                "phone": phone,
            },
        )

    # --------------------------------------------------------
    # RESPONSE
    # --------------------------------------------------------

    def _send_json(
        self,
        status_code: int,
        data: dict,
    ) -> None:
        body = json.dumps(
            data,
            ensure_ascii=False,
        ).encode("utf-8")

        self.send_response(
            status_code
        )

        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8",
        )

        self.send_header(
            "Content-Length",
            str(len(body)),
        )

        # Очень важно для HTTP/1.1-клиентов.
        self.send_header(
            "Connection",
            "close",
        )

        self.end_headers()

        try:
            self.wfile.write(body)
            self.wfile.flush()

        except Exception:
            logger.exception(
                "HTTP: failed to send response"
            )


# ============================================================
# SERVER START
# ============================================================


def start_ipc_server(
    state_machine: CallStateMachine,
    host: str = "127.0.0.1",
    port: int = 8765,
) -> ThreadingHTTPServer:
    """
    Start local IPC HTTP server.

    MicroSIP -> HTTP -> Queue -> Worker -> StateMachine
    """

    worker = IPCEventWorker(
        state_machine
    )

    worker.start()

    IPCHandler.event_worker = worker

    server = ThreadingHTTPServer(
        (host, port),
        IPCHandler,
    )

    # Сохраняем worker на сервере, чтобы его можно было
    # корректно остановить при shutdown.
    server.ipc_worker = worker

    logger.info(
        "IPC server started on %s:%s",
        host,
        port,
    )

    return server