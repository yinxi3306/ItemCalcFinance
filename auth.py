from __future__ import annotations

import hashlib
import secrets
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Literal

ROLE_USER: Literal["user"] = "user"
ROLE_ADMIN: Literal["admin"] = "admin"

PBKDF2_ITERATIONS = 210_000

_USERS_SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT COLLATE NOCASE NOT NULL UNIQUE,
    salt TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK(role IN ('user', 'admin')),
    created_at TEXT NOT NULL
);
"""


@dataclass(frozen=True)
class UserSession:
    username: str
    role: str

    @property
    def is_admin(self) -> bool:
        return self.role == ROLE_ADMIN


def ensure_users_table(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.executescript(_USERS_SCHEMA)
        conn.commit()


def _hash_password(password: str) -> tuple[str, str]:
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("ascii"),
        PBKDF2_ITERATIONS,
    )
    return salt, dk.hex()


def _verify_password(password: str, salt: str, stored_hex: str) -> bool:
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("ascii"),
        PBKDF2_ITERATIONS,
    )
    return secrets.compare_digest(dk.hex(), stored_hex)


def _validate_username(username: str) -> str:
    u = username.strip()
    if len(u) < 2:
        raise ValueError("用户名至少 2 个字符。")
    if len(u) > 64:
        raise ValueError("用户名不能超过 64 个字符。")
    return u


def _validate_password(password: str) -> None:
    if len(password) < 6:
        raise ValueError("密码至少 6 位。")
    if len(password) > 256:
        raise ValueError("密码过长。")


def register_user(db_path: Path, username: str, password: str) -> str:
    u = _validate_username(username)
    _validate_password(password)
    ensure_users_table(db_path)
    salt, phash = _hash_password(password)
    created = datetime.now().replace(microsecond=0).isoformat(sep=" ")
    with sqlite3.connect(db_path) as conn:
        n = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        role = ROLE_ADMIN if n == 0 else ROLE_USER
        try:
            conn.execute(
                """
                INSERT INTO users (username, salt, password_hash, role, created_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (u, salt, phash, role, created),
            )
            conn.commit()
        except sqlite3.IntegrityError as e:
            raise ValueError("该用户名已被注册。") from e
    return role


def authenticate(db_path: Path, username: str, password: str) -> UserSession | None:
    ensure_users_table(db_path)
    u = username.strip()
    if not u:
        return None
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT username, salt, password_hash, role FROM users WHERE username = ?",
            (u,),
        ).fetchone()
    if row is None:
        return None
    stored_user, salt, phash, role = row
    if not _verify_password(password, salt, phash):
        return None
    return UserSession(username=stored_user, role=role)
