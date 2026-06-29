"""FastAPI application entry point."""

from contextlib import asynccontextmanager

from fastapi import FastAPI

from app.database import init_db
from app.routes import router


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield


app = FastAPI(
    title="Lustia API Search",
    description="Lustia API for uploading and full-text searching databases (<1 s)",
    version="0.1.0",
    lifespan=lifespan,
)

app.include_router(router, prefix="/lustia/api/search")


@app.get("/health")
async def health():
    return {"status": "ok"}
