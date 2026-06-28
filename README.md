# DB Search API

Fast API for uploading and searching databases with sub-second response times.

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
| POST | `/api/v1/databases/upload?db_name=X` | Upload a CSV/JSON file |
| GET | `/api/v1/databases` | List all databases |
| DELETE | `/api/v1/databases/{db_name}` | Delete a database |
| GET | `/api/v1/search?q=term` | Full-text search |

## Running tests

```bash
pytest
```
