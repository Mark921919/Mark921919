"""Shared test fixtures."""

from pathlib import Path

import pytest

from app.database import init_db, reset_db, set_db_path


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path: Path):
    """Each test gets its own fresh SQLite database."""
    db_path = tmp_path / "test.db"
    set_db_path(db_path)
    init_db()
    yield
    reset_db()
