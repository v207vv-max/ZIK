from __future__ import annotations

import json
import logging
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from .calls.state_machine import CallStateMachine


logger = logging.getLogger("ZIK.IPC")


class IPCHandler(BaseHTTPRequestHandler):

    state_machine: CallStateMachine | None = None

    def log_message(self, format: str, *args) -> None:
        logger.info(
            "HTTP | " + format,
            *args,
        )

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

        content_length = int(
            self.headers.get("Content-Length", "0")
        )

        raw_body = self.rfile.read(content_length)

        try:
            data = json.loads(
                raw_body.decode("utf-8")
            )

        except (UnicodeDecodeError, json.JSONDecodeError):
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": "Invalid JSON",
                },
            )
            return

        event_type = data.get("event")
        phone = data.get("phone")

        try:

            if event_type == "outgoing":
                call_id = self.state_machine.outgoing(phone)

                result = {
                    "ok": True,
                    "call_id": call_id,
                }

            elif event_type == "start":
                self.state_machine.start(phone)

                result = {
                    "ok": True,
                }

            elif event_type == "busy":
                self.state_machine.busy(phone)

                result = {
                    "ok": True,
                }

            elif event_type == "end":
                self.state_machine.end(phone)

                result = {
                    "ok": True,
                }

            elif event_type == "kochadan":
                self.state_machine.kochadan()

                result = {
                    "ok": True,
                }

            else:
                self._send_json(
                    400,
                    {
                        "ok": False,
                        "error": f"Unknown event: {event_type}",
                    },
                )
                return

            self._send_json(200, result)

        except Exception as exc:
            logger.exception(
                "IPC event failed"
            )

            self._send_json(
                500,
                {
                    "ok": False,
                    "error": str(exc),
                },
            )

    def _send_json(
        self,
        status_code: int,
        data: dict,
    ) -> None:

        payload = json.dumps(
            data,
            ensure_ascii=False,
        ).encode("utf-8")

        self.send_response(status_code)

        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8",
        )

        self.send_header(
            "Content-Length",
            str(len(payload)),
        )

        self.end_headers()

        self.wfile.write(payload)


def start_ipc_server(
    state_machine: CallStateMachine,
) -> ThreadingHTTPServer:

    IPCHandler.state_machine = state_machine

    server = ThreadingHTTPServer(
        ("127.0.0.1", 8765),
        IPCHandler,
    )

    logger.info(
        "IPC server started on 127.0.0.1:8765"
    )

    return server