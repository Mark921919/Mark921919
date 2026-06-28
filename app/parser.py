"""Parse uploaded files (CSV, JSON) into flat text rows for indexing."""

import csv
import io
import json


def parse_csv(content: bytes, encoding: str = "utf-8") -> list[str]:
    text = content.decode(encoding)
    reader = csv.DictReader(io.StringIO(text))
    rows: list[str] = []
    for record in reader:
        row_text = " | ".join(f"{k}: {v}" for k, v in record.items() if v)
        if row_text:
            rows.append(row_text)
    return rows


def parse_json(content: bytes, encoding: str = "utf-8") -> list[str]:
    text = content.decode(encoding)
    data = json.loads(text)

    if isinstance(data, list):
        items = data
    elif isinstance(data, dict):
        items = [data]
    else:
        raise ValueError(f"Unsupported JSON root type: {type(data).__name__}")

    rows: list[str] = []
    for item in items:
        if isinstance(item, dict):
            row_text = " | ".join(f"{k}: {v}" for k, v in item.items() if v is not None)
        else:
            row_text = str(item)
        if row_text:
            rows.append(row_text)
    return rows


def parse_file(filename: str, content: bytes) -> list[str]:
    lower = filename.lower()
    if lower.endswith(".csv"):
        return parse_csv(content)
    if lower.endswith(".json"):
        return parse_json(content)
    raise ValueError(f"Unsupported file type: {filename}. Use .csv or .json")
