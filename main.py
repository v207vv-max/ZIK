from __future__ import annotations

import logging
import threading
from pathlib import Path

from app.calls.state_machine import CallStateMachine
from app.config import load_config
from app.controller import ZIKController
from app.ipc_server import start_ipc_server
from app.logger import setup_logging
from app.microsip.controller import MicroSIPController
from app.storage.database import Database
from app.ui.main_window import MainWindow


# ============================================================
# DEFAULT EXCEL
# ============================================================

EXCEL_FILE = Path(
    r"C:\Users\Asko\Desktop\call-sotuv\temur\Новая папка\Sverka.xlsx"
)


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    setup_logging()

    logger = logging.getLogger("ZIK")

    logger.info("========================================")
    logger.info("ZIK starting...")
    logger.info("========================================")

    # ========================================================
    # CONFIG
    # ========================================================

    config = load_config()

    # ========================================================
    # MICROSIP
    # ========================================================

    microsip = MicroSIPController(
        config.microsip.executable
    )

    microsip.validate()

    logger.info(
        "MicroSIP validated: %s",
        config.microsip.executable,
    )

    # ========================================================
    # DATABASE
    # ========================================================

    database = Database()

    logger.info(
        "SQLite: %s",
        database.path,
    )

    # ========================================================
    # STATE MACHINE
    # ========================================================

    state_machine = CallStateMachine(
        database=database,
        microsip=microsip,
        ring_timeout_seconds=(
            config.microsip.ring_timeout_seconds
        ),
    )

    logger.info(
        "StateMachine initialized"
    )

    # ========================================================
    # CONTROLLER
    # ========================================================

    controller = ZIKController(
        microsip=microsip,
        state_machine=state_machine,
    )

    logger.info(
        "Controller initialized"
    )

    # ========================================================
    # DEFAULT EXCEL
    # ========================================================

    if EXCEL_FILE.exists():
        try:
            controller.load_excel(
                EXCEL_FILE
            )

            logger.info(
                "Excel loaded: %s",
                EXCEL_FILE,
            )

        except Exception:
            logger.exception(
                "Failed to load default Excel"
            )

    else:
        logger.warning(
            "Default Excel not found: %s",
            EXCEL_FILE,
        )

    # ========================================================
    # IPC SERVER
    # ========================================================

    server = None
    ipc_thread: threading.Thread | None = None

    try:
        server = start_ipc_server(
            state_machine
        )

        # ----------------------------------------------------
        # IMPORTANT
        #
        # HTTP server MUST run in a separate thread.
        #
        # Tkinter UI and IPC must work simultaneously.
        # ----------------------------------------------------

        ipc_thread = threading.Thread(
            target=server.serve_forever,
            name="ZIK-IPC-Server",
            daemon=True,
        )

        ipc_thread.start()

        logger.info(
            "IPC server thread started"
        )

    except Exception:
        logger.exception(
            "Failed to start IPC server"
        )

        try:
            controller.stop()
        except Exception:
            logger.exception(
                "Failed to stop controller "
                "after IPC startup error"
            )

        raise

    # ========================================================
    # CORE READY
    # ========================================================

    logger.info("========================================")
    logger.info("ZIK CORE READY")
    logger.info(
        "MicroSIP: %s",
        config.microsip.executable,
    )
    logger.info(
        "Watchdog: %s sec",
        config.microsip.ring_timeout_seconds,
    )
    logger.info(
        "SQLite: %s",
        database.path,
    )
    logger.info(
        "Excel: %s",
        EXCEL_FILE,
    )
    logger.info(
        "IPC: 127.0.0.1:8765",
    )
    logger.info("========================================")

    # ========================================================
    # UI
    # ========================================================

    app = None

    try:
        app = MainWindow(
            controller=controller,
        )

        logger.info(
            "MainWindow initialized"
        )

        # ----------------------------------------------------
        # Tkinter main loop
        #
        # IPC continues working in its own thread.
        # ----------------------------------------------------

        app.run()

    except KeyboardInterrupt:
        logger.info(
            "ZIK: KeyboardInterrupt"
        )

    except Exception:
        logger.exception(
            "ZIK: unexpected UI error"
        )

    finally:
        # ====================================================
        # SHUTDOWN
        # ====================================================

        logger.info(
            "ZIK: shutting down..."
        )

        # ----------------------------------------------------
        # 1. Stop Controller / Dialer
        # ----------------------------------------------------

        try:
            controller.stop()

        except Exception:
            logger.exception(
                "ZIK: failed to stop controller"
            )

        # ----------------------------------------------------
        # 2. Stop IPC server
        # ----------------------------------------------------

        if server is not None:

            try:
                server.shutdown()

                logger.info(
                    "ZIK: IPC server shutdown requested"
                )

            except Exception:
                logger.exception(
                    "ZIK: failed to shutdown IPC server"
                )

            # ------------------------------------------------
            # Wait for IPC server thread
            # ------------------------------------------------

            if (
                ipc_thread is not None
                and ipc_thread.is_alive()
            ):
                try:
                    ipc_thread.join(
                        timeout=2.0
                    )

                    logger.info(
                        "ZIK: IPC server thread stopped"
                    )

                except Exception:
                    logger.exception(
                        "ZIK: failed to join IPC thread"
                    )

            # ------------------------------------------------
            # Stop IPC worker
            # ------------------------------------------------

            ipc_worker = getattr(
                server,
                "ipc_worker",
                None,
            )

            if ipc_worker is not None:
                try:
                    ipc_worker.stop()

                    logger.info(
                        "ZIK: IPC worker stopped"
                    )

                except Exception:
                    logger.exception(
                        "ZIK: failed to stop IPC worker"
                    )

            # ------------------------------------------------
            # Close server socket
            # ------------------------------------------------

            try:
                server.server_close()

                logger.info(
                    "ZIK: IPC server socket closed"
                )

            except Exception:
                logger.exception(
                    "ZIK: failed to close IPC server"
                )

        # ----------------------------------------------------
        # 3. StateMachine shutdown
        # ----------------------------------------------------

        try:
            state_machine.shutdown()

        except Exception:
            logger.exception(
                "ZIK: failed to shutdown StateMachine"
            )

        # ----------------------------------------------------
        # DONE
        # ----------------------------------------------------

        logger.info("========================================")
        logger.info("ZIK STOPPED")
        logger.info("========================================")


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()