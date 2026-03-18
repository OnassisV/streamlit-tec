from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
import json
import os
import sqlite3

import pandas as pd


class StorageBackend(ABC):
    mode = "none"

    @abstractmethod
    def save_run(self, payload: dict) -> None:
        raise NotImplementedError

    @abstractmethod
    def list_recent_runs(self, limit: int = 20) -> pd.DataFrame:
        raise NotImplementedError


class NullStorageBackend(StorageBackend):
    mode = "none"

    def save_run(self, payload: dict) -> None:
        return None

    def list_recent_runs(self, limit: int = 20) -> pd.DataFrame:
        return pd.DataFrame(
            columns=[
                "processed_at",
                "source_name",
                "input_rows",
                "clean_rows",
                "deleted_rows",
                "pending_rows",
                "storage_mode",
            ]
        )


class SQLiteStorageBackend(StorageBackend):
    mode = "sqlite"

    def __init__(self, db_path: str) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_schema()

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.db_path)

    def _ensure_schema(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS app_runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    processed_at TEXT NOT NULL,
                    source_name TEXT,
                    input_rows INTEGER,
                    clean_rows INTEGER,
                    deleted_rows INTEGER,
                    pending_rows INTEGER,
                    config_json TEXT,
                    mapping_json TEXT,
                    notes_json TEXT
                )
                """
            )

    def save_run(self, payload: dict) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO app_runs (
                    processed_at,
                    source_name,
                    input_rows,
                    clean_rows,
                    deleted_rows,
                    pending_rows,
                    config_json,
                    mapping_json,
                    notes_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    payload.get("processed_at") or datetime.now().isoformat(timespec="seconds"),
                    payload.get("source_name"),
                    payload.get("input_rows"),
                    payload.get("clean_rows"),
                    payload.get("deleted_rows"),
                    payload.get("pending_rows"),
                    json.dumps(payload.get("config", {}), ensure_ascii=False),
                    json.dumps(payload.get("mapping", {}), ensure_ascii=False),
                    json.dumps(payload.get("notes", {}), ensure_ascii=False),
                ),
            )

    def list_recent_runs(self, limit: int = 20) -> pd.DataFrame:
        query = """
            SELECT
                processed_at,
                source_name,
                input_rows,
                clean_rows,
                deleted_rows,
                pending_rows
            FROM app_runs
            ORDER BY id DESC
            LIMIT ?
        """
        with self._connect() as conn:
            df = pd.read_sql_query(query, conn, params=(limit,))
        df["storage_mode"] = self.mode
        return df


def build_storage_backend(secrets: dict | None = None) -> StorageBackend:
    secrets = secrets or {}
    mode = str(secrets.get("APP_STORAGE_MODE") or os.getenv("APP_STORAGE_MODE", "none")).strip().lower()

    if mode in {"", "none", "null"}:
        return NullStorageBackend()

    if mode == "sqlite":
        db_path = str(
            secrets.get("APP_SQLITE_PATH")
            or os.getenv("APP_SQLITE_PATH", "data/app_runs.db")
        ).strip()
        return SQLiteStorageBackend(db_path)

    return NullStorageBackend()
