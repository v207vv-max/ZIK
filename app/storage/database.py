from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[2]

DATA_DIR = PROJECT_ROOT / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

DATABASE_FILE = DATA_DIR / "zik.db"


class Database:
    def __init__(self, path: Path = DATABASE_FILE) -> None:
        self.path = path

        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(
            self.path,
            timeout=10,
        )

        connection.row_factory = sqlite3.Row

        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS calls (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,

                    phone TEXT NOT NULL,

                    state TEXT NOT NULL,
                    result TEXT NOT NULL,

                    outgoing_at TEXT,
                    started_at TEXT,
                    ended_at TEXT,

                    manual_result TEXT,

                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )

            connection.commit()

    def create_call(
        self,
        phone: str,
        state: str,
        result: str,
        outgoing_at: datetime,
    ) -> int:

        now = datetime.now().isoformat()

        with self._connect() as connection:
            cursor = connection.execute(
                """
                INSERT INTO calls (
                    phone,
                    state,
                    result,
                    outgoing_at,
                    created_at,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    phone,
                    state,
                    result,
                    outgoing_at.isoformat(),
                    now,
                    now,
                ),
            )

            connection.commit()

            return int(cursor.lastrowid)

    def update_call(
        self,
        call_id: int,
        *,
        state: str | None = None,
        result: str | None = None,
        started_at: datetime | None = None,
        ended_at: datetime | None = None,
        manual_result: str | None = None,
    ) -> None:

        fields: list[str] = []
        values: list[Any] = []

        if state is not None:
            fields.append("state = ?")
            values.append(state)

        if result is not None:
            fields.append("result = ?")
            values.append(result)

        if started_at is not None:
            fields.append("started_at = ?")
            values.append(started_at.isoformat())

        if ended_at is not None:
            fields.append("ended_at = ?")
            values.append(ended_at.isoformat())

        if manual_result is not None:
            fields.append("manual_result = ?")
            values.append(manual_result)

        fields.append("updated_at = ?")
        values.append(datetime.now().isoformat())

        values.append(call_id)

        query = f"""
            UPDATE calls
            SET {", ".join(fields)}
            WHERE id = ?
        """

        with self._connect() as connection:
            connection.execute(query, values)
            connection.commit()

    def get_active_call(self) -> sqlite3.Row | None:
        with self._connect() as connection:
            cursor = connection.execute(
                """
                SELECT *
                FROM calls
                WHERE state IN ('OUTGOING', 'IN_CALL')
                ORDER BY id DESC
                LIMIT 1
                """
            )

            return cursor.fetchone()

    def get_call(self, call_id: int) -> sqlite3.Row | None:
        with self._connect() as connection:
            cursor = connection.execute(
                """
                SELECT *
                FROM calls
                WHERE id = ?
                """,
                (call_id,),
            )

            return cursor.fetchone()