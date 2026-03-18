from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
import json
import os
import sqlite3

import pandas as pd

from app_auth import ROLE_DEFINITIONS, hash_password, normalize_username, serialize_user_record, verify_password

try:
    import pymysql
    from pymysql.cursors import DictCursor
except ImportError:
    pymysql = None
    DictCursor = None


class StorageBackend(ABC):
    mode = "none"
    supports_auth = False
    supports_user_management = False

    @abstractmethod
    def save_run(self, payload: dict) -> None:
        raise NotImplementedError

    @abstractmethod
    def list_recent_runs(self, limit: int = 20) -> pd.DataFrame:
        raise NotImplementedError

    def auth_enabled(self) -> bool:
        return self.supports_auth

    def has_users(self) -> bool:
        return False

    def authenticate_user(self, username: str, password: str) -> dict:
        return {"ok": False, "reason": "auth_not_available"}

    def create_initial_admin(
        self,
        username: str,
        full_name: str,
        password: str,
        email: str | None = None,
        phone_number: str | None = None,
    ) -> None:
        raise NotImplementedError("Authentication is not available for this backend.")

    def list_users(self) -> pd.DataFrame:
        return pd.DataFrame(
            columns=[
                "id",
                "username",
                "full_name",
                "email",
                "phone_number",
                "role_key",
                "role_name",
                "is_enabled",
                "active_from",
                "active_until",
                "last_login_at",
                "created_at",
            ]
        )

    def create_user(self, payload: dict) -> None:
        raise NotImplementedError("User management is not available for this backend.")

    def update_user(self, user_id: int, payload: dict) -> None:
        raise NotImplementedError("User management is not available for this backend.")


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


class MySQLStorageBackend(StorageBackend):
    mode = "mysql"
    supports_auth = True
    supports_user_management = True

    def __init__(
        self,
        host: str,
        port: int,
        database: str,
        user: str,
        password: str,
        charset: str = "utf8mb4",
        connect_timeout: int = 10,
        ssl_disabled: bool = False,
        ssl_ca: str | None = None,
        ssl_cert: str | None = None,
        ssl_key: str | None = None,
        ssl_verify_cert: bool | None = None,
        ssl_verify_identity: bool | None = None,
    ) -> None:
        if pymysql is None:
            raise RuntimeError("Falta la dependencia `pymysql`. Agregala a requirements.txt.")

        self.host = host
        self.port = port
        self.database = database
        self.user = user
        self.password = password
        self.charset = charset
        self.connect_timeout = connect_timeout
        self.ssl_disabled = ssl_disabled
        self.ssl_ca = ssl_ca
        self.ssl_cert = ssl_cert
        self.ssl_key = ssl_key
        self.ssl_verify_cert = ssl_verify_cert
        self.ssl_verify_identity = ssl_verify_identity
        self._ensure_schema()

    def _connect(self):
        connection_kwargs = {
            "host": self.host,
            "port": self.port,
            "user": self.user,
            "password": self.password,
            "database": self.database,
            "charset": self.charset,
            "cursorclass": DictCursor,
            "autocommit": False,
            "connect_timeout": self.connect_timeout,
        }
        if self.ssl_disabled:
            connection_kwargs["ssl_disabled"] = True
        else:
            if self.ssl_ca:
                connection_kwargs["ssl_ca"] = self.ssl_ca
            if self.ssl_cert:
                connection_kwargs["ssl_cert"] = self.ssl_cert
            if self.ssl_key:
                connection_kwargs["ssl_key"] = self.ssl_key
            if self.ssl_verify_cert is not None:
                connection_kwargs["ssl_verify_cert"] = self.ssl_verify_cert
            if self.ssl_verify_identity is not None:
                connection_kwargs["ssl_verify_identity"] = self.ssl_verify_identity
        return pymysql.connect(**connection_kwargs)

    def _query_dataframe(self, query: str, params=None, columns: list[str] | None = None) -> pd.DataFrame:
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params or ())
                rows = cursor.fetchall()

        if not rows:
            return pd.DataFrame(columns=columns)

        return pd.DataFrame(rows, columns=columns)

    def _ensure_schema(self) -> None:
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS app_runs (
                        id BIGINT PRIMARY KEY AUTO_INCREMENT,
                        processed_at DATETIME NOT NULL,
                        source_name VARCHAR(255),
                        input_rows INT,
                        clean_rows INT,
                        deleted_rows INT,
                        pending_rows INT,
                        config_json TEXT,
                        mapping_json TEXT,
                        notes_json TEXT
                    )
                    """
                )
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS app_roles (
                        role_key VARCHAR(50) PRIMARY KEY,
                        role_name VARCHAR(100) NOT NULL,
                        description TEXT,
                        permissions_json TEXT NOT NULL
                    )
                    """
                )
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS app_users (
                        id BIGINT PRIMARY KEY AUTO_INCREMENT,
                        username VARCHAR(64) NOT NULL UNIQUE,
                        full_name VARCHAR(255) NOT NULL,
                        email VARCHAR(255) NOT NULL UNIQUE,
                        phone_number VARCHAR(32) NOT NULL,
                        password_hash VARCHAR(128) NOT NULL,
                        password_salt VARCHAR(64) NOT NULL,
                        role_key VARCHAR(50) NOT NULL,
                        is_enabled BOOLEAN NOT NULL DEFAULT TRUE,
                        active_from DATETIME NULL,
                        active_until DATETIME NULL,
                        last_login_at DATETIME NULL,
                        created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                        updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        created_by BIGINT NULL,
                        CONSTRAINT fk_app_users_role FOREIGN KEY (role_key) REFERENCES app_roles(role_key)
                    )
                    """
                )
            self._migrate_app_users_schema(conn)
            self._seed_roles(conn)
            conn.commit()

    def _column_metadata(self, conn, table_name: str, column_name: str) -> dict | None:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT
                    COLUMN_NAME,
                    IS_NULLABLE,
                    COLUMN_TYPE
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = %s
                  AND TABLE_NAME = %s
                  AND COLUMN_NAME = %s
                LIMIT 1
                """,
                (self.database, table_name, column_name),
            )
            return cursor.fetchone()

    def _migrate_app_users_schema(self, conn) -> None:
        email_metadata = self._column_metadata(conn, "app_users", "email")
        phone_metadata = self._column_metadata(conn, "app_users", "phone_number")

        with conn.cursor() as cursor:
            if not phone_metadata:
                cursor.execute(
                    "ALTER TABLE app_users ADD COLUMN phone_number VARCHAR(32) NULL AFTER email"
                )
                phone_metadata = self._column_metadata(conn, "app_users", "phone_number")

            cursor.execute(
                """
                SELECT COUNT(*) AS total
                FROM app_users
                WHERE email IS NULL
                   OR TRIM(email) = ''
                   OR phone_number IS NULL
                   OR TRIM(phone_number) = ''
                """
            )
            invalid_rows = int(cursor.fetchone()["total"])

            if invalid_rows == 0:
                if email_metadata and email_metadata["IS_NULLABLE"] == "YES":
                    cursor.execute(
                        "ALTER TABLE app_users MODIFY COLUMN email VARCHAR(255) NOT NULL"
                    )
                if phone_metadata and phone_metadata["IS_NULLABLE"] == "YES":
                    cursor.execute(
                        "ALTER TABLE app_users MODIFY COLUMN phone_number VARCHAR(32) NOT NULL"
                    )

    def _seed_roles(self, conn) -> None:
        with conn.cursor() as cursor:
            for role_key, role_data in ROLE_DEFINITIONS.items():
                cursor.execute(
                    """
                    INSERT INTO app_roles (role_key, role_name, description, permissions_json)
                    VALUES (%s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                        role_name = VALUES(role_name),
                        description = VALUES(description),
                        permissions_json = VALUES(permissions_json)
                    """,
                    (
                        role_key,
                        role_data["label"],
                        role_data["description"],
                        json.dumps(role_data["permissions"], ensure_ascii=False),
                    ),
                )

    def save_run(self, payload: dict) -> None:
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
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
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        payload.get("processed_at") or datetime.now().replace(microsecond=0),
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
            conn.commit()

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
            LIMIT %s
        """
        df = self._query_dataframe(
            query,
            params=(limit,),
            columns=[
                "processed_at",
                "source_name",
                "input_rows",
                "clean_rows",
                "deleted_rows",
                "pending_rows",
            ],
        )
        df["storage_mode"] = self.mode
        return df

    def has_users(self) -> bool:
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) AS total FROM app_users")
                row = cursor.fetchone()
        return bool(row and row["total"] > 0)

    def authenticate_user(self, username: str, password: str) -> dict:
        normalized_username = normalize_username(username)
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    """
                    SELECT
                        u.id,
                        u.username,
                        u.full_name,
                        u.email,
                        u.phone_number,
                        u.password_hash,
                        u.password_salt,
                        u.role_key,
                        u.is_enabled,
                        u.active_from,
                        u.active_until,
                        u.last_login_at,
                        r.role_name,
                        r.permissions_json
                    FROM app_users u
                    INNER JOIN app_roles r ON r.role_key = u.role_key
                    WHERE u.username = %s
                    LIMIT 1
                    """,
                    (normalized_username,),
                )
                record = cursor.fetchone()
                if not record:
                    return {"ok": False, "reason": "invalid_credentials"}
                if not verify_password(password, record["password_salt"], record["password_hash"]):
                    return {"ok": False, "reason": "invalid_credentials"}

                now = datetime.now()
                if not bool(record["is_enabled"]):
                    return {"ok": False, "reason": "disabled"}
                if record["active_from"] and now < record["active_from"]:
                    return {"ok": False, "reason": "not_yet_active"}
                if record["active_until"] and now > record["active_until"]:
                    return {"ok": False, "reason": "expired"}

                cursor.execute(
                    "UPDATE app_users SET last_login_at = %s WHERE id = %s",
                    (now, record["id"]),
                )
            conn.commit()

        return {"ok": True, "user": serialize_user_record(record)}

    def create_initial_admin(
        self,
        username: str,
        full_name: str,
        password: str,
        email: str | None = None,
        phone_number: str | None = None,
    ) -> None:
        if self.has_users():
            raise ValueError("Ya existe al menos un usuario. Usa la gestion de usuarios para crear mas cuentas.")
        self.create_user(
            {
                "username": username,
                "full_name": full_name,
                "email": email,
                "phone_number": phone_number,
                "password": password,
                "role_key": "super_admin",
                "is_enabled": True,
                "active_from": None,
                "active_until": None,
                "created_by": None,
            }
        )

    def list_users(self) -> pd.DataFrame:
        query = """
            SELECT
                u.id,
                u.username,
                u.full_name,
                u.email,
                u.phone_number,
                u.role_key,
                r.role_name,
                u.is_enabled,
                u.active_from,
                u.active_until,
                u.last_login_at,
                u.created_at
            FROM app_users u
            INNER JOIN app_roles r ON r.role_key = u.role_key
            ORDER BY u.username ASC
        """
        return self._query_dataframe(
            query,
            columns=[
                "id",
                "username",
                "full_name",
                "email",
                "phone_number",
                "role_key",
                "role_name",
                "is_enabled",
                "active_from",
                "active_until",
                "last_login_at",
                "created_at",
            ],
        )

    def create_user(self, payload: dict) -> None:
        username = normalize_username(payload.get("username") or "")
        full_name = str(payload.get("full_name") or "").strip()
        password = str(payload.get("password") or "")
        role_key = str(payload.get("role_key") or "").strip()
        email = str(payload.get("email") or "").strip()
        phone_number = str(payload.get("phone_number") or "").strip()

        if not username or not full_name or not password or role_key not in ROLE_DEFINITIONS:
            raise ValueError("Usuario, nombre completo, contrasena y rol son obligatorios.")
        if not email or not phone_number:
            raise ValueError("Correo electronico y celular son obligatorios.")

        salt_hex, hash_hex = hash_password(password)
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    """
                    INSERT INTO app_users (
                        username,
                        full_name,
                        email,
                        phone_number,
                        password_hash,
                        password_salt,
                        role_key,
                        is_enabled,
                        active_from,
                        active_until,
                        created_by
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        username,
                        full_name,
                        email,
                        phone_number,
                        hash_hex,
                        salt_hex,
                        role_key,
                        bool(payload.get("is_enabled", True)),
                        payload.get("active_from"),
                        payload.get("active_until"),
                        payload.get("created_by"),
                    ),
                )
            conn.commit()

    def update_user(self, user_id: int, payload: dict) -> None:
        fields = []
        params = []

        simple_fields = {
            "full_name": payload.get("full_name"),
            "email": (str(payload.get("email") or "").strip() or None) if "email" in payload else None,
            "phone_number": (str(payload.get("phone_number") or "").strip() or None) if "phone_number" in payload else None,
            "role_key": payload.get("role_key"),
            "is_enabled": payload.get("is_enabled"),
            "active_from": payload.get("active_from"),
            "active_until": payload.get("active_until"),
        }
        for field_name, field_value in simple_fields.items():
            if field_name not in payload:
                continue
            if field_name == "role_key" and field_value not in ROLE_DEFINITIONS:
                raise ValueError("Rol no valido.")
            if field_name in {"email", "phone_number"} and not field_value:
                raise ValueError("Correo electronico y celular son obligatorios.")
            fields.append(f"{field_name} = %s")
            params.append(field_value)

        password = str(payload.get("password") or "")
        if password:
            salt_hex, hash_hex = hash_password(password)
            fields.extend(["password_hash = %s", "password_salt = %s"])
            params.extend([hash_hex, salt_hex])

        if not fields:
            return

        params.append(user_id)
        with self._connect() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    f"UPDATE app_users SET {', '.join(fields)} WHERE id = %s",
                    tuple(params),
                )
            conn.commit()


def build_storage_backend(secrets: dict | None = None) -> StorageBackend:
    secrets = secrets or {}
    mode = str(secrets.get("APP_STORAGE_MODE") or os.getenv("APP_STORAGE_MODE", "none")).strip().lower()

    def get_bool_setting(secret_key: str, env_key: str, default: bool | None = None) -> bool | None:
        raw_value = secrets.get(secret_key)
        if raw_value is None:
            raw_value = os.getenv(env_key)
        if raw_value is None:
            return default
        return str(raw_value).strip().lower() in {"1", "true", "yes", "on"}

    if mode in {"", "none", "null"}:
        return NullStorageBackend()

    if mode == "sqlite":
        db_path = str(
            secrets.get("APP_SQLITE_PATH")
            or os.getenv("APP_SQLITE_PATH", "data/app_runs.db")
        ).strip()
        return SQLiteStorageBackend(db_path)

    if mode == "mysql":
        host = str(secrets.get("APP_MYSQL_HOST") or os.getenv("APP_MYSQL_HOST", "127.0.0.1")).strip()
        port = int(secrets.get("APP_MYSQL_PORT") or os.getenv("APP_MYSQL_PORT", "3306"))
        database = str(secrets.get("APP_MYSQL_DATABASE") or os.getenv("APP_MYSQL_DATABASE", "")).strip()
        user = str(secrets.get("APP_MYSQL_USER") or os.getenv("APP_MYSQL_USER", "")).strip()
        password = str(secrets.get("APP_MYSQL_PASSWORD") or os.getenv("APP_MYSQL_PASSWORD", "")).strip()
        charset = str(secrets.get("APP_MYSQL_CHARSET") or os.getenv("APP_MYSQL_CHARSET", "utf8mb4")).strip()
        connect_timeout = int(secrets.get("APP_MYSQL_CONNECT_TIMEOUT") or os.getenv("APP_MYSQL_CONNECT_TIMEOUT", "10"))
        ssl_disabled = bool(get_bool_setting("APP_MYSQL_SSL_DISABLED", "APP_MYSQL_SSL_DISABLED", default=False))
        ssl_ca = str(secrets.get("APP_MYSQL_SSL_CA") or os.getenv("APP_MYSQL_SSL_CA", "")).strip() or None
        ssl_cert = str(secrets.get("APP_MYSQL_SSL_CERT") or os.getenv("APP_MYSQL_SSL_CERT", "")).strip() or None
        ssl_key = str(secrets.get("APP_MYSQL_SSL_KEY") or os.getenv("APP_MYSQL_SSL_KEY", "")).strip() or None
        ssl_verify_cert = get_bool_setting("APP_MYSQL_SSL_VERIFY_CERT", "APP_MYSQL_SSL_VERIFY_CERT", default=None)
        ssl_verify_identity = get_bool_setting("APP_MYSQL_SSL_VERIFY_IDENTITY", "APP_MYSQL_SSL_VERIFY_IDENTITY", default=None)
        if not all([host, database, user, password]):
            raise RuntimeError("Faltan variables de entorno o secrets de MySQL para inicializar la app.")
        return MySQLStorageBackend(
            host=host,
            port=port,
            database=database,
            user=user,
            password=password,
            charset=charset,
            connect_timeout=connect_timeout,
            ssl_disabled=ssl_disabled,
            ssl_ca=ssl_ca,
            ssl_cert=ssl_cert,
            ssl_key=ssl_key,
            ssl_verify_cert=ssl_verify_cert,
            ssl_verify_identity=ssl_verify_identity,
        )

    return NullStorageBackend()
