"""Tests for API routes (integration tests)."""

import io
import json

import pytest
from httpx import ASGITransport, AsyncClient

from app.main import app


@pytest.fixture
async def client():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as ac:
        yield ac


class TestHealthEndpoint:
    @pytest.mark.asyncio
    async def test_health(self, client):
        resp = await client.get("/health")
        assert resp.status_code == 200
        assert resp.json()["status"] == "ok"


class TestUploadEndpoint:
    @pytest.mark.asyncio
    async def test_upload_csv(self, client):
        csv_data = b"name,age\nAlice,30\nBob,25\n"
        resp = await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "people"},
            files={"file": ("people.csv", io.BytesIO(csv_data), "text/csv")},
        )
        assert resp.status_code == 200
        body = resp.json()
        assert body["status"] == "ok"
        assert body["records_inserted"] == 2
        assert body["elapsed_seconds"] < 1.0

    @pytest.mark.asyncio
    async def test_upload_json(self, client):
        data = [{"color": "red"}, {"color": "blue"}]
        resp = await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "colors"},
            files={
                "file": ("colors.json", io.BytesIO(json.dumps(data).encode()), "application/json")
            },
        )
        assert resp.status_code == 200
        assert resp.json()["records_inserted"] == 2

    @pytest.mark.asyncio
    async def test_upload_duplicate_name(self, client):
        csv_data = b"a,b\n1,2\n"
        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "dup"},
            files={"file": ("d.csv", io.BytesIO(csv_data), "text/csv")},
        )
        resp = await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "dup"},
            files={"file": ("d.csv", io.BytesIO(csv_data), "text/csv")},
        )
        assert resp.status_code == 409

    @pytest.mark.asyncio
    async def test_upload_empty_file(self, client):
        resp = await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "empty"},
            files={"file": ("e.csv", io.BytesIO(b""), "text/csv")},
        )
        assert resp.status_code == 400

    @pytest.mark.asyncio
    async def test_upload_unsupported_type(self, client):
        resp = await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "bad"},
            files={"file": ("data.xml", io.BytesIO(b"<xml/>"), "text/xml")},
        )
        assert resp.status_code == 400


class TestListDatabases:
    @pytest.mark.asyncio
    async def test_list_empty(self, client):
        resp = await client.get("/api/v1/databases")
        assert resp.status_code == 200
        assert resp.json()["databases"] == []

    @pytest.mark.asyncio
    async def test_list_after_upload(self, client):
        csv_data = b"x,y\n1,2\n"
        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "listed"},
            files={"file": ("f.csv", io.BytesIO(csv_data), "text/csv")},
        )
        resp = await client.get("/api/v1/databases")
        dbs = resp.json()["databases"]
        assert len(dbs) >= 1
        names = [d.get("db_name") or d.get("name") for d in dbs]
        assert "listed" in names


class TestDeleteEndpoint:
    @pytest.mark.asyncio
    async def test_delete_existing(self, client):
        csv_data = b"a,b\n1,2\n"
        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "del_me"},
            files={"file": ("f.csv", io.BytesIO(csv_data), "text/csv")},
        )
        resp = await client.delete("/api/v1/databases/del_me")
        assert resp.status_code == 200
        assert resp.json()["status"] == "deleted"

    @pytest.mark.asyncio
    async def test_delete_nonexistent(self, client):
        resp = await client.delete("/api/v1/databases/nope")
        assert resp.status_code == 404


class TestSearchEndpoint:
    @pytest.mark.asyncio
    async def test_search_basic(self, client):
        csv_data = b"name,city\nAlice,Moscow\nBob,Berlin\n"
        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "cities"},
            files={"file": ("c.csv", io.BytesIO(csv_data), "text/csv")},
        )
        resp = await client.get("/api/v1/search", params={"q": "Moscow"})
        assert resp.status_code == 200
        body = resp.json()
        assert body["count"] >= 1
        assert body["elapsed_seconds"] < 1.0

    @pytest.mark.asyncio
    async def test_search_with_db_filter(self, client):
        csv1 = b"val\nfoo\n"
        csv2 = b"val\nfoo\n"
        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "f1"},
            files={"file": ("a.csv", io.BytesIO(csv1), "text/csv")},
        )
        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "f2"},
            files={"file": ("b.csv", io.BytesIO(csv2), "text/csv")},
        )
        resp = await client.get("/api/v1/search", params={"q": "foo", "db_name": "f1"})
        assert resp.status_code == 200
        assert resp.json()["count"] == 1

    @pytest.mark.asyncio
    async def test_search_nonexistent_db(self, client):
        resp = await client.get("/api/v1/search", params={"q": "x", "db_name": "nope"})
        assert resp.status_code == 404

    @pytest.mark.asyncio
    async def test_search_empty_query(self, client):
        resp = await client.get("/api/v1/search", params={"q": ""})
        assert resp.status_code == 422

    @pytest.mark.asyncio
    async def test_search_performance(self, client):
        """Upload 500 records and verify search < 1s."""
        import csv as csv_mod

        buf = io.StringIO()
        writer = csv_mod.writer(buf)
        writer.writerow(["id", "text"])
        for i in range(500):
            writer.writerow([i, f"data entry record number {i} with extra text"])
        csv_bytes = buf.getvalue().encode()

        await client.post(
            "/api/v1/databases/upload",
            params={"db_name": "perf"},
            files={"file": ("perf.csv", io.BytesIO(csv_bytes), "text/csv")},
        )
        resp = await client.get("/api/v1/search", params={"q": "record"})
        assert resp.status_code == 200
        assert resp.json()["elapsed_seconds"] < 1.0
