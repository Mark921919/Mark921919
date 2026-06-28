"""API route definitions."""

import time

from fastapi import APIRouter, HTTPException, Query, UploadFile

from app.database import (
    count_records,
    create_database_entry,
    delete_database_entry,
    get_database_entry,
    insert_records,
    list_databases,
    search_records,
)
from app.parser import SUPPORTED_EXTENSIONS, parse_file

router = APIRouter()


@router.post("/databases/upload")
async def upload_database(file: UploadFile, db_name: str = Query(..., min_length=1)):
    """Upload a file as a new searchable database (CSV, TSV, JSON, XML, Excel, TXT)."""
    start = time.perf_counter()

    if not file.filename:
        raise HTTPException(status_code=400, detail="File has no name")

    existing = get_database_entry(db_name)
    if existing:
        raise HTTPException(status_code=409, detail=f"Database '{db_name}' already exists")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Empty file")

    try:
        rows = parse_file(file.filename, content)
    except (ValueError, UnicodeDecodeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    if not rows:
        raise HTTPException(status_code=400, detail="File contains no records")

    db_id = create_database_entry(db_name)
    inserted = insert_records(db_id, rows)
    elapsed = time.perf_counter() - start

    return {
        "status": "ok",
        "db_name": db_name,
        "records_inserted": inserted,
        "elapsed_seconds": round(elapsed, 4),
    }


@router.get("/databases")
async def list_all_databases():
    """List all uploaded databases."""
    dbs = list_databases()
    result = []
    for db in dbs:
        cnt = count_records(db["id"])
        result.append({**db, "record_count": cnt})
    return {"databases": result}


@router.delete("/databases/{db_name}")
async def delete_database(db_name: str):
    """Delete a database and all its records."""
    entry = get_database_entry(db_name)
    if not entry:
        raise HTTPException(status_code=404, detail=f"Database '{db_name}' not found")
    delete_database_entry(entry["id"])
    return {"status": "deleted", "db_name": db_name}


@router.get("/formats")
async def supported_formats():
    """List all supported file formats for upload."""
    return {"formats": SUPPORTED_EXTENSIONS}


@router.get("/search")
async def search(
    q: str = Query(..., min_length=1, description="Search query"),
    db_name: str | None = Query(None, description="Limit to specific database"),
    limit: int = Query(50, ge=1, le=500),
):
    """Full-text search across all databases (or a specific one)."""
    start = time.perf_counter()

    if db_name:
        entry = get_database_entry(db_name)
        if not entry:
            raise HTTPException(status_code=404, detail=f"Database '{db_name}' not found")

    try:
        results = search_records(q, limit=limit, db_name=db_name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Search error: {exc}")

    elapsed = time.perf_counter() - start

    return {
        "query": q,
        "db_name": db_name,
        "count": len(results),
        "elapsed_seconds": round(elapsed, 4),
        "results": results,
    }
