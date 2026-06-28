"""Tests for app.parser module."""

import json

import pytest

from app.parser import parse_csv, parse_file, parse_json


class TestParseCsv:
    def test_basic_csv(self):
        content = b"name,age,city\nAlice,30,NYC\nBob,25,LA\n"
        rows = parse_csv(content)
        assert len(rows) == 2
        assert "name: Alice" in rows[0]
        assert "age: 30" in rows[0]
        assert "city: NYC" in rows[0]

    def test_empty_values_skipped(self):
        content = b"a,b\n1,\n"
        rows = parse_csv(content)
        assert len(rows) == 1
        assert "b:" not in rows[0]

    def test_headers_only(self):
        content = b"a,b,c\n"
        rows = parse_csv(content)
        assert rows == []

    def test_utf8_content(self):
        content = "name,city\nИван,Москва\n".encode("utf-8")
        rows = parse_csv(content)
        assert len(rows) == 1
        assert "Иван" in rows[0]
        assert "Москва" in rows[0]


class TestParseJson:
    def test_list_of_dicts(self):
        data = [{"name": "Alice", "age": 30}, {"name": "Bob", "age": 25}]
        rows = parse_json(json.dumps(data).encode())
        assert len(rows) == 2
        assert "name: Alice" in rows[0]

    def test_single_dict(self):
        data = {"key": "value"}
        rows = parse_json(json.dumps(data).encode())
        assert len(rows) == 1
        assert "key: value" in rows[0]

    def test_none_values_skipped(self):
        data = [{"a": 1, "b": None}]
        rows = parse_json(json.dumps(data).encode())
        assert len(rows) == 1
        assert "b:" not in rows[0]

    def test_list_of_primitives(self):
        data = ["hello", "world"]
        rows = parse_json(json.dumps(data).encode())
        assert rows == ["hello", "world"]

    def test_unsupported_root_type(self):
        with pytest.raises(ValueError, match="Unsupported JSON root type"):
            parse_json(b'"just a string"')

    def test_empty_list(self):
        rows = parse_json(b"[]")
        assert rows == []


class TestParseFile:
    def test_csv_extension(self):
        rows = parse_file("data.csv", b"x,y\n1,2\n")
        assert len(rows) == 1

    def test_json_extension(self):
        rows = parse_file("data.json", json.dumps([{"a": 1}]).encode())
        assert len(rows) == 1

    def test_uppercase_extension(self):
        rows = parse_file("DATA.CSV", b"x,y\n1,2\n")
        assert len(rows) == 1

    def test_unsupported_extension(self):
        with pytest.raises(ValueError, match="Unsupported file type"):
            parse_file("data.xml", b"<root/>")
