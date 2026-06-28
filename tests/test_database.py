"""Tests for app.database module."""

import pytest

from app.database import (
    count_records,
    create_database_entry,
    delete_database_entry,
    get_database_entry,
    insert_records,
    list_databases,
    search_records,
)


class TestDatabaseEntry:
    def test_create_and_get(self):
        db_id = create_database_entry("test_db")
        assert db_id > 0
        entry = get_database_entry("test_db")
        assert entry is not None
        assert entry["name"] == "test_db"
        assert entry["id"] == db_id

    def test_get_nonexistent(self):
        assert get_database_entry("nope") is None

    def test_create_duplicate_raises(self):
        create_database_entry("dup")
        with pytest.raises(Exception):
            create_database_entry("dup")

    def test_list_databases(self):
        create_database_entry("a")
        create_database_entry("b")
        dbs = list_databases()
        names = [d["name"] for d in dbs]
        assert "a" in names
        assert "b" in names

    def test_delete_database(self):
        db_id = create_database_entry("to_delete")
        insert_records(db_id, ["row1", "row2"])
        delete_database_entry(db_id)
        assert get_database_entry("to_delete") is None
        assert count_records(db_id) == 0


class TestRecords:
    def test_insert_and_count(self):
        db_id = create_database_entry("rec_db")
        inserted = insert_records(db_id, ["a", "b", "c"])
        assert inserted == 3
        assert count_records(db_id) == 3

    def test_insert_empty(self):
        db_id = create_database_entry("empty_db")
        inserted = insert_records(db_id, [])
        assert inserted == 0
        assert count_records(db_id) == 0


class TestSearch:
    def test_basic_search(self):
        db_id = create_database_entry("search_db")
        insert_records(
            db_id,
            [
                "name: Alice | city: New York",
                "name: Bob | city: Los Angeles",
                "name: Charlie | city: Chicago",
            ],
        )
        results = search_records("Alice")
        assert len(results) == 1
        assert "Alice" in results[0]["row_data"]

    def test_search_no_results(self):
        db_id = create_database_entry("s_db")
        insert_records(db_id, ["hello world"])
        results = search_records("zzzznotfound")
        assert results == []

    def test_search_with_db_filter(self):
        db1 = create_database_entry("db_one")
        db2 = create_database_entry("db_two")
        insert_records(db1, ["common term alpha"])
        insert_records(db2, ["common term beta"])
        results = search_records("common", db_name="db_one")
        assert len(results) == 1
        assert results[0]["db_name"] == "db_one"

    def test_search_limit(self):
        db_id = create_database_entry("limit_db")
        insert_records(db_id, [f"item number {i}" for i in range(100)])
        results = search_records("item", limit=5)
        assert len(results) == 5

    def test_search_across_databases(self):
        db1 = create_database_entry("multi1")
        db2 = create_database_entry("multi2")
        insert_records(db1, ["shared keyword here"])
        insert_records(db2, ["shared keyword there"])
        results = search_records("shared")
        assert len(results) == 2

    def test_search_performance_small(self):
        """Verify search completes in well under 1 second for small datasets."""
        import time

        db_id = create_database_entry("perf_db")
        insert_records(db_id, [f"record data entry number {i} info" for i in range(1000)])
        start = time.perf_counter()
        results = search_records("record")
        elapsed = time.perf_counter() - start
        assert elapsed < 1.0
        assert len(results) > 0
