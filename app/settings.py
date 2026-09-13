from __future__ import annotations

import json
from pathlib import Path


class Settings:
    def __init__(self, path: Path = Path("config/settings.json")) -> None:
        self.path = path

        self.data = {
            "last_excel_file": "",
            "last_excel_row": 2,
            "wps_sync": True,
        }

        self.load()

    def load(self) -> None:
        if not self.path.exists():
            return

        try:
            self.data.update(
                json.loads(
                    self.path.read_text(encoding="utf-8")
                )
            )
        except Exception:
            pass

    def save(self) -> None:
        self.path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.path.write_text(
            json.dumps(
                self.data,
                ensure_ascii=False,
                indent=4,
            ),
            encoding="utf-8",
        )

    @property
    def last_excel_file(self) -> str:
        return self.data["last_excel_file"]

    @last_excel_file.setter
    def last_excel_file(self, value: str) -> None:
        self.data["last_excel_file"] = value

    @property
    def last_excel_row(self) -> int:
        return int(self.data["last_excel_row"])

    @last_excel_row.setter
    def last_excel_row(self, value: int) -> None:
        self.data["last_excel_row"] = int(value)

    @property
    def wps_sync(self) -> bool:
        return bool(self.data["wps_sync"])

    @wps_sync.setter
    def wps_sync(self, value: bool) -> None:
        self.data["wps_sync"] = bool(value)