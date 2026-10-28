"""FastAPI 本地网页/API。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from .service import WorkflowService


class CorrectionRequest(BaseModel):
    reviewed_by: str = Field(min_length=1)
    reason: str = Field(min_length=1)
    corrected_revenue: dict[str, str | None]


def create_app(data_root: Path) -> FastAPI:
    service = WorkflowService(data_root)
    app = FastAPI(title="仁恒合并利润表工作台", version="1.0.0")
    app.state.workflow_service = service
    web_dir = Path(__file__).with_name("web")
    app.mount("/static", StaticFiles(directory=web_dir), name="static")

    @app.on_event("shutdown")
    def shutdown() -> None:
        service.close()

    @app.get("/", response_class=HTMLResponse)
    def index() -> HTMLResponse:
        return HTMLResponse((web_dir / "index.html").read_text(encoding="utf-8"))

    @app.get("/api/runs")
    def list_runs() -> list[dict[str, Any]]:
        return service.list_runs()

    @app.post("/api/runs")
    async def create_run(
        year: int = Form(...),
        month: int = Form(...),
        operator: str = Form(...),
        files: list[UploadFile] = File(...),
    ) -> dict[str, Any]:
        try:
            return service.create_run(year=year, month=month, operator=operator, uploads=files)
        except (ValueError, OSError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error
        finally:
            for upload in files:
                await upload.close()

    @app.get("/api/runs/{run_id}")
    def get_run(run_id: str) -> dict[str, Any]:
        try:
            return service.get_status(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error

    @app.post("/api/runs/{run_id}/files")
    async def add_files(
        run_id: str,
        files: list[UploadFile] = File(...),
    ) -> dict[str, Any]:
        try:
            return service.add_uploads(run_id, files)
        except (KeyError, ValueError, OSError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error
        finally:
            for upload in files:
                await upload.close()

    @app.post("/api/runs/{run_id}/preflight")
    def preflight_run(run_id: str) -> dict[str, Any]:
        try:
            return service.preflight(run_id)
        except (KeyError, ValueError, RuntimeError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error

    @app.post("/api/runs/{run_id}/start")
    def start_run(run_id: str) -> dict[str, Any]:
        try:
            return service.start(run_id)
        except (KeyError, ValueError, RuntimeError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error

    @app.post("/api/runs/{run_id}/resume")
    def resume_run(run_id: str, request: CorrectionRequest) -> dict[str, Any]:
        try:
            return service.resume(run_id, request.model_dump(mode="json"))
        except (KeyError, ValueError, RuntimeError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error

    @app.get("/api/runs/{run_id}/artifacts/{key}")
    def download_artifact(run_id: str, key: str) -> FileResponse:
        try:
            path = service.artifact_path(run_id, key)
        except (KeyError, FileNotFoundError) as error:
            raise HTTPException(status_code=404, detail=str(error)) from error
        return FileResponse(path, filename=path.name)

    return app
