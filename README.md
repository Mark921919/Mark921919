# Lustia API Search

Lustia API for uploading and searching databases with sub-second response times.

## Features

- Upload CSV/JSON files as searchable databases
- Full-text search powered by SQLite FTS5
- Upload and search in < 1 second
- Filter search by specific database

## Quick start

```bash
pip install -e ".[dev]"
uvicorn app.main:app --reload
```

API docs: http://localhost:8000/docs

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/lustia/api/search/databases/upload?db_name=X` | Upload a CSV/JSON file |
| GET | `/lustia/api/search/databases` | List all databases |
| DELETE | `/lustia/api/search/databases/{db_name}` | Delete a database |
| GET | `/lustia/api/search/search?q=term` | Full-text search |

## Running tests

```bash
pytest
```
