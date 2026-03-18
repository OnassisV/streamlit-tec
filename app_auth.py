from __future__ import annotations

from datetime import datetime
import hashlib
import hmac
import json
import os

import streamlit as st


SESSION_USER_KEY = "auth_user"
PBKDF2_ITERATIONS = 210_000
PASSWORD_SALT_BYTES = 16

ROLE_DEFINITIONS = {
    "super_admin": {
        "label": "Administrador general",
        "description": "Control total de configuraciones, usuarios y vigencias.",
        "permissions": [
            "process_files",
            "manage_general_config",
            "manage_users",
            "manage_user_activation",
            "view_history",
        ],
    },
    "admin_operaciones": {
        "label": "Administrador operativo",
        "description": "Puede procesar, revisar historial y habilitar o deshabilitar usuarios por fechas.",
        "permissions": [
            "process_files",
            "manage_user_activation",
            "view_history",
        ],
    },
    "analista": {
        "label": "Analista",
        "description": "Puede usar el procesamiento con la configuracion general definida.",
        "permissions": [
            "process_files",
            "view_history",
        ],
    },
}


def normalize_username(username: str) -> str:
    return username.strip().lower()


def hash_password(password: str, salt_hex: str | None = None) -> tuple[str, str]:
    salt = bytes.fromhex(salt_hex) if salt_hex else os.urandom(PASSWORD_SALT_BYTES)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PBKDF2_ITERATIONS,
    )
    return salt.hex(), digest.hex()


def verify_password(password: str, salt_hex: str, expected_hash_hex: str) -> bool:
    _, calculated_hash = hash_password(password, salt_hex=salt_hex)
    return hmac.compare_digest(calculated_hash, expected_hash_hex)


def get_role_permissions(role_key: str) -> list[str]:
    role = ROLE_DEFINITIONS.get(role_key, {})
    permissions = role.get("permissions", [])
    return [str(permission) for permission in permissions]


def serialize_user_record(record: dict) -> dict:
    permissions_raw = record.get("permissions_json") or "[]"
    try:
        permissions = json.loads(permissions_raw)
    except json.JSONDecodeError:
        permissions = []

    return {
        "id": record.get("id"),
        "username": record.get("username"),
        "full_name": record.get("full_name"),
        "email": record.get("email"),
        "phone_number": record.get("phone_number"),
        "role_key": record.get("role_key"),
        "role_label": record.get("role_name") or ROLE_DEFINITIONS.get(record.get("role_key"), {}).get("label"),
        "permissions": [str(permission) for permission in permissions],
        "is_enabled": bool(record.get("is_enabled")),
        "active_from": record.get("active_from"),
        "active_until": record.get("active_until"),
        "last_login_at": record.get("last_login_at"),
    }


def get_authenticated_user() -> dict | None:
    return st.session_state.get(SESSION_USER_KEY)


def set_authenticated_user(user: dict) -> None:
    st.session_state[SESSION_USER_KEY] = user


def clear_authenticated_user() -> None:
    st.session_state.pop(SESSION_USER_KEY, None)


def user_has_permission(permission: str, user: dict | None = None) -> bool:
    current_user = user or get_authenticated_user()
    if not current_user:
        return False
    permissions = current_user.get("permissions") or []
    return permission in permissions


def describe_access_window(active_from, active_until) -> str:
    if not active_from and not active_until:
        return "Sin fecha de vencimiento"
    if active_from and active_until:
        return f"Activo desde {active_from:%Y-%m-%d %H:%M} hasta {active_until:%Y-%m-%d %H:%M}"
    if active_from:
        return f"Activo desde {active_from:%Y-%m-%d %H:%M}"
    return f"Activo hasta {active_until:%Y-%m-%d %H:%M}"