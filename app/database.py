"""SQLite database layer with FTS5 full-text search."""

import sqlite3
import threading
from pathlib import Path

_DB_DIR = Path(__file__).resolve().parent.parent / "data"
_DB_PATH = _DB_DIR / "store.db"

_local = threading.local()


def _get_conn() -> sqlite3.Connection:
    conn = getattr(_local, "conn", None)
    if conn is None:
        _DB_DIR.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(_DB_PATH), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        _local.conn = conn
    return conn


def init_db() -> None:
    conn = _get_conn()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS databases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            db_id INTEGER NOT NULL REFERENCES databases(id) ON DELETE CASCADE,
            row_data TEXT NOT NULL
        );

        CREATE VIRTUAL TABLE IF NOT EXISTS records_fts USING fts5(
            row_data,
            content='records',
            content_rowid='id'
        );

        CREATE TRIGGER IF NOT EXISTS records_ai AFTER INSERT ON records BEGIN
            INSERT INTO records_fts(rowid, row_data) VALUES (new.id, new.row_data);
        END;

        CREATE TRIGGER IF NOT EXISTS records_ad AFTER DELETE ON records BEGIN
            INSERT INTO records_fts(records_fts, rowid, row_data)
                VALUES ('delete', old.id, old.row_data);
        END;
        """
    )
    conn.commit()


def reset_db() -> None:
    """Drop all tables and reinitialise. Used by tests."""
    conn = _get_conn()
    conn.executescript(
        """
        DROP TRIGGER IF EXISTS records_ai;
        DROP TRIGGER IF EXISTS records_ad;
        DROP TABLE IF EXISTS records_fts;
        DROP TABLE IF EXISTS records;
        DROP TABLE IF EXISTS databases;
        """
    )
    conn.commit()
    init_db()


def create_database_entry(name: str) -> int:
    conn = _get_conn()
    cur = conn.execute("INSERT INTO databases (name) VALUES (?)", (name,))
    conn.commit()
    return cur.lastrowid  # type: ignore[return-value]


def get_database_entry(name: str) -> dict | None:
    conn = _get_conn()
    row = conn.execute("SELECT * FROM databases WHERE name = ?", (name,)).fetchone()
    return dict(row) if row else None


def list_databases() -> list[dict]:
    conn = _get_conn()
    rows = conn.execute("SELECT * FROM databases ORDER BY id").fetchall()
    return [dict(r) for r in rows]


def delete_database_entry(db_id: int) -> None:
    conn = _get_conn()
    conn.execute("DELETE FROM records WHERE db_id = ?", (db_id,))
    conn.execute("DELETE FROM databases WHERE id = ?", (db_id,))
    conn.commit()


def insert_records(db_id: int, rows: list[str]) -> int:
    conn = _get_conn()
    conn.executemany(
        "INSERT INTO records (db_id, row_data) VALUES (?, ?)",
        [(db_id, r) for r in rows],
    )
    conn.commit()
    return len(rows)


def search_records(query: str, limit: int = 50, db_name: str | None = None) -> list[dict]:
    conn = _get_conn()
    if db_name:
        sql = """
            SELECT r.id, d.name AS db_name, r.row_data,
                   rank
            FROM records_fts fts
            JOIN records r ON r.id = fts.rowid
            JOIN databases d ON d.id = r.db_id
            WHERE fts.row_data MATCH ? AND d.name = ?
            ORDER BY rank
            LIMIT ?
        """
        rows = conn.execute(sql, (query, db_name, limit)).fetchall()
    else:
        sql = """
            SELECT r.id, d.name AS db_name, r.row_data,
                   rank
            FROM records_fts fts
            JOIN records r ON r.id = fts.rowid
            JOIN databases d ON d.id = r.db_id
            WHERE fts.row_data MATCH ?
            ORDER BY rank
            LIMIT ?
        """
        rows = conn.execute(sql, (query, limit)).fetchall()
    return [dict(r) for r in rows]


def count_records(db_id: int) -> int:
    conn = _get_conn()
    row = conn.execute("SELECT COUNT(*) AS cnt FROM records WHERE db_id = ?", (db_id,)).fetchone()
    return row["cnt"]  # type: ignore[index]


def set_db_path(path: Path) -> None:
    """Override the database path (for tests)."""
    global _DB_PATH, _DB_DIR
    _DB_PATH = path
    _DB_DIR = path.parent
    # Reset any existing connection
    _local.conn = None
