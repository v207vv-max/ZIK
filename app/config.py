from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_FILE = PROJECT_ROOT / "config" / "config.json"


@dataclass(frozen=True)
class MicroSIPConfig:
    executable: Path
    ring_timeout_seconds: int
    max_call_duration_seconds: int


@dataclass(frozen=True)
class HotkeyConfig:
    kucha: str


@dataclass(frozen=True)
class Config:
    microsip: MicroSIPConfig
    hotkeys: HotkeyConfig


def load_config() -> Config:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {CONFIG_FILE}"
        )

    with CONFIG_FILE.open("r", encoding="utf-8") as file:
        data = json.load(file)

    microsip_data = data["microsip"]
    hotkeys_data = data["hotkeys"]

    return Config(
        microsip=MicroSIPConfig(
            executable=Path(microsip_data["executable"]),
            ring_timeout_seconds=int(
                microsip_data["ring_timeout_seconds"]
            ),
            max_call_duration_seconds=int(
                microsip_data["max_call_duration_seconds"]
            ),
        ),
        hotkeys=HotkeyConfig(
            kucha=hotkeys_data["kucha"],
        ),
    )