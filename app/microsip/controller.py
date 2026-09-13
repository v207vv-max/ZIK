from __future__ import annotations

import logging
import subprocess
from pathlib import Path


logger = logging.getLogger(__name__)


class MicroSIPError(RuntimeError):
    """Base exception for MicroSIP integration errors."""


class MicroSIPController:
    def __init__(self, executable: Path) -> None:
        self.executable = executable

    def validate(self) -> None:
        if not self.executable.exists():
            raise MicroSIPError(
                f"MicroSIP executable not found: {self.executable}"
            )

        if self.executable.suffix.lower() != ".exe":
            raise MicroSIPError(
                f"Expected a Windows executable: {self.executable}"
            )

    def call(self, number: str) -> None:
        self.validate()

        number = number.strip()

        if not number:
            raise ValueError("Phone number is empty")

        logger.info("Starting call: %s", number)

        subprocess.Popen(
            [
                str(self.executable),
                number,
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

    def hangup_all(self) -> None:
        self.validate()

        logger.info("Hanging up all MicroSIP calls")

        subprocess.Popen(
            [
                str(self.executable),
                "/hangupall",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )