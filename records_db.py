from __future__ import annotations

import sqlite3
from pathlib import Path

_SCHEMA = """
CREATE TABLE IF NOT EXISTS line_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_name TEXT NOT NULL,
    quantity TEXT NOT NULL,
    unit_price TEXT NOT NULL,
    total_price TEXT NOT NULL,
    submitted_at TEXT NOT NULL,
    submitted_by TEXT NOT NULL DEFAULT ''
);
"""


def ensure_schema(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.executescript(_SCHEMA)
        cols = {row[1] for row in conn.execute("PRAGMA table_info(line_items)")}
        if "submitted_by" not in cols:
            conn.execute(
                "ALTER TABLE line_items ADD COLUMN submitted_by TEXT NOT NULL DEFAULT ''"
            )
        conn.commit()


def insert_line_items_batch(
    db_path: Path,
    rows: list[tuple[str, str, str, str, str, str]],
) -> None:
    """Insert many line items in one transaction.

    Each tuple is (product_name, quantity, unit_price, total_price, submitted_at, submitted_by).
    """
    if not rows:
        return
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        try:
            conn.executemany(
                """
                INSERT INTO line_items
                    (product_name, quantity, unit_price, total_price, submitted_at, submitted_by)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                rows,
            )
            conn.commit()
        except sqlite3.Error:
            conn.rollback()
            raise


def delete_line_items_for_submitter(db_path: Path, username: str) -> int:
    """Remove all line items submitted by the given username. Returns rows deleted."""
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute(
            "DELETE FROM line_items WHERE submitted_by = ?",
            (username,),
        )
        conn.commit()
        return int(cur.rowcount) if cur.rowcount is not None else 0


def delete_all_line_items(db_path: Path) -> int:
    """Remove all line items. Returns rows deleted."""
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute("DELETE FROM line_items")
        conn.commit()
        return int(cur.rowcount) if cur.rowcount is not None else 0


def insert_line_item(
    db_path: Path,
    *,
    product_name: str,
    quantity: str,
    unit_price: str,
    total_price: str,
    submitted_at: str,
    submitted_by: str,
) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO line_items
                (product_name, quantity, unit_price, total_price, submitted_at, submitted_by)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (product_name, quantity, unit_price, total_price, submitted_at, submitted_by),
        )
        conn.commit()
