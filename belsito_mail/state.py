"""Persist only opaque completion IDs and sample-test acknowledgement locally."""

from __future__ import annotations

import os
from pathlib import Path
import sqlite3


def default_state_path() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    if base:
        return Path(base) / "BelsitoPlumbing" / "EmailAssistant" / "state.sqlite3"
    return Path.home() / ".local" / "share" / "belsito-email-assistant" / "state.sqlite3"


class LocalState:
    """The database never receives message bodies, subjects, addresses or passwords."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(self.path)
        self.connection.execute(
            "CREATE TABLE IF NOT EXISTS completed ("
            "source TEXT NOT NULL, item_id TEXT NOT NULL, "
            "completed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, "
            "PRIMARY KEY (source, item_id))"
        )
        self.connection.execute(
            "CREATE TABLE IF NOT EXISTS settings (name TEXT PRIMARY KEY, value TEXT NOT NULL)"
        )
        self.connection.commit()

    def completed_ids(self, source: str) -> set[str]:
        return {row[0] for row in self.connection.execute(
            "SELECT item_id FROM completed WHERE source = ?", (source,)
        )}

    def set_completed(self, source: str, item_id: str, completed: bool) -> None:
        with self.connection:
            if completed:
                self.connection.execute(
                    "INSERT OR IGNORE INTO completed(source, item_id) VALUES (?, ?)",
                    (source, item_id),
                )
            else:
                self.connection.execute(
                    "DELETE FROM completed WHERE source = ? AND item_id = ?", (source, item_id)
                )

    @property
    def sample_test_passed(self) -> bool:
        row = self.connection.execute(
            "SELECT value FROM settings WHERE name = 'sample_test_passed'"
        ).fetchone()
        return bool(row and row[0] == "yes")

    def acknowledge_sample_test(self) -> None:
        with self.connection:
            self.connection.execute(
                "INSERT OR REPLACE INTO settings(name, value) VALUES ('sample_test_passed', 'yes')"
            )

    def close(self) -> None:
        self.connection.close()
