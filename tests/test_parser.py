"""Tests for app.parser module."""

import io
import json

import pytest
from openpyxl import Workbook

from app.parser import (
    SUPPORTED_EXTENSIONS,
    parse_csv,
    parse_excel,
    parse_file,
    parse_json,
    parse_tsv,
    parse_txt,
    parse_xml,
)


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


class TestParseTsv:
    def test_basic_tsv(self):
        content = b"name\tage\nAlice\t30\nBob\t25\n"
        rows = parse_tsv(content)
        assert len(rows) == 2
        assert "name: Alice" in rows[0]
        assert "age: 30" in rows[0]

    def test_empty_values_skipped(self):
        content = b"a\tb\n1\t\n"
        rows = parse_tsv(content)
        assert len(rows) == 1
        assert "b:" not in rows[0]

    def test_headers_only(self):
        content = b"a\tb\tc\n"
        rows = parse_tsv(content)
        assert rows == []


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


class TestParseXml:
    def test_basic_xml(self):
        xml = b"""<root>
            <item><name>Alice</name><age>30</age></item>
            <item><name>Bob</name><age>25</age></item>
        </root>"""
        rows = parse_xml(xml)
        assert len(rows) == 2
        assert "name: Alice" in rows[0]
        assert "age: 30" in rows[0]

    def test_xml_with_attributes(self):
        xml = b"""<root>
            <item id="1"><name>Alice</name></item>
        </root>"""
        rows = parse_xml(xml)
        assert len(rows) == 1
        assert "id: 1" in rows[0]
        assert "name: Alice" in rows[0]

    def test_empty_xml(self):
        xml = b"<root></root>"
        rows = parse_xml(xml)
        assert rows == []

    def test_xml_text_in_element(self):
        xml = b"<root><item>direct text</item></root>"
        rows = parse_xml(xml)
        assert len(rows) == 1
        assert "item: direct text" in rows[0]


class TestParseExcel:
    @staticmethod
    def _make_xlsx(headers: list[str], data: list[list]) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for row in data:
            ws.append(row)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_basic_excel(self):
        content = self._make_xlsx(["name", "age"], [["Alice", 30], ["Bob", 25]])
        rows = parse_excel(content)
        assert len(rows) == 2
        assert "name: Alice" in rows[0]
        assert "age: 30" in rows[0]

    def test_excel_skip_empty_cells(self):
        content = self._make_xlsx(["a", "b"], [[1, None]])
        rows = parse_excel(content)
        assert len(rows) == 1
        assert "b:" not in rows[0]

    def test_excel_headers_only(self):
        content = self._make_xlsx(["a", "b"], [])
        rows = parse_excel(content)
        assert rows == []


class TestParseTxt:
    def test_basic_txt(self):
        content = b"line one\nline two\nline three\n"
        rows = parse_txt(content)
        assert len(rows) == 3
        assert rows[0] == "line one"

    def test_empty_lines_skipped(self):
        content = b"hello\n\n\nworld\n"
        rows = parse_txt(content)
        assert len(rows) == 2
        assert rows == ["hello", "world"]

    def test_whitespace_stripped(self):
        content = b"  spaced  \n"
        rows = parse_txt(content)
        assert rows == ["spaced"]

    def test_empty_file(self):
        rows = parse_txt(b"")
        assert rows == []


class TestParseFile:
    def test_csv_extension(self):
        rows = parse_file("data.csv", b"x,y\n1,2\n")
        assert len(rows) == 1

    def test_tsv_extension(self):
        rows = parse_file("data.tsv", b"x\ty\n1\t2\n")
        assert len(rows) == 1

    def test_json_extension(self):
        rows = parse_file("data.json", json.dumps([{"a": 1}]).encode())
        assert len(rows) == 1

    def test_xml_extension(self):
        rows = parse_file("data.xml", b"<r><i><a>1</a></i></r>")
        assert len(rows) == 1

    def test_xlsx_extension(self):
        content = TestParseExcel._make_xlsx(["h"], [["v"]])
        rows = parse_file("data.xlsx", content)
        assert len(rows) == 1

    def test_txt_extension(self):
        rows = parse_file("data.txt", b"hello\n")
        assert len(rows) == 1

    def test_uppercase_extension(self):
        rows = parse_file("DATA.CSV", b"x,y\n1,2\n")
        assert len(rows) == 1

    def test_unsupported_extension(self):
        with pytest.raises(ValueError, match="Unsupported file type"):
            parse_file("data.pdf", b"content")

    def test_supported_extensions_dict(self):
        assert ".csv" in SUPPORTED_EXTENSIONS
        assert ".tsv" in SUPPORTED_EXTENSIONS
        assert ".json" in SUPPORTED_EXTENSIONS
        assert ".xml" in SUPPORTED_EXTENSIONS
        assert ".xlsx" in SUPPORTED_EXTENSIONS
        assert ".txt" in SUPPORTED_EXTENSIONS
