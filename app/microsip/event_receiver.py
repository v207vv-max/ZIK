from __future__ import annotations

import json
import logging
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]

LOG_DIR = PROJECT_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOG_DIR / "microsip_events.log"

ZIK_ENDPOINT = "http://127.0.0.1:8765/event"


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | "
            "%(levelname)s | "
            "%(message)s"
        ),
        handlers=[
            logging.FileHandler(
                LOG_FILE,
                encoding="utf-8",
            ),
            logging.StreamHandler(),
        ],
    )


def send_to_zik(
    event_type: str,
    phone: str | None,
) -> None:

    payload = {
        "event": event_type,
        "phone": phone,
        "timestamp": datetime.now(
            timezone.utc
        ).isoformat(),
    }

    data = json.dumps(
        payload
    ).encode("utf-8")

    request = urllib.request.Request(
        ZIK_ENDPOINT,
        data=data,
        headers={
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(
            request,
            timeout=2,
        ) as response:

            response_body = response.read().decode(
                "utf-8"
            )

            logging.getLogger(
                "ZIK.MicroSIP"
            ).info(
                "ZIK response: %s",
                response_body,
            )

    except (
        urllib.error.URLError,
        TimeoutError,
        OSError,
    ) as exc:

        logging.getLogger(
            "ZIK.MicroSIP"
        ).error(
            "Failed to send event to ZIK: %s",
            exc,
        )


def main() -> None:

    setup_logging()

    logger = logging.getLogger(
        "ZIK.MicroSIP"
    )

    arguments = sys.argv[1:]

    if not arguments:
        event_type = "unknown"
        phone = None

    else:
        event_type = arguments[0]
        phone = (
            arguments[1]
            if len(arguments) > 1
            else None
        )

    logger.info(
        "EVENT | type=%s | phone=%s",
        event_type,
        phone,
    )

    if event_type in {
        "outgoing",
        "start",
        "busy",
        "end",
    }:
        send_to_zik(
            event_type,
            phone,
        )


if __name__ == "__main__":
    main()