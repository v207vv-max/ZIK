from __future__ import annotations

import logging

from app.calls.state_machine import CallStateMachine
from app.config import load_config
from app.ipc_server import start_ipc_server
from app.logger import setup_logging
from app.microsip.controller import MicroSIPController
from app.storage.database import Database


def main() -> None:
    setup_logging()

    logger = logging.getLogger("ZIK")

    config = load_config()

    microsip = MicroSIPController(
        config.microsip.executable
    )

    microsip.validate()

    database = Database()

    state_machine = CallStateMachine(
        database=database,
        microsip=microsip,
        ring_timeout_seconds=config.microsip.ring_timeout_seconds,
    )

    server = start_ipc_server(state_machine)

    logger.info("========================================")
    logger.info("ZIK CORE")
    logger.info("MicroSIP: %s", config.microsip.executable)
    logger.info(
        "Ring timeout: %s sec",
        config.microsip.ring_timeout_seconds,
    )
    logger.info("SQLite: %s", database.path)
    logger.info("IPC: http://127.0.0.1:8765")
    logger.info("ZIK is ready")
    logger.info("========================================")

    try:
        server.serve_forever()

    except KeyboardInterrupt:
        logger.info("Stopping ZIK...")

    finally:
        server.server_close()


if __name__ == "__main__":
    main()