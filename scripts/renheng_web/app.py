"""FastAPI网页和REST API。用户身份来自签名会话，不接受客户端自报审核人。"""
from __future__ import annotations

from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any, Callable
from urllib.parse import quote

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from .security import SESSION_COOKIE, SessionTokens
from .service import WorkflowService


class LoginRequest(BaseModel):
    username: str = Field(min_length=1, max_length=80)
    password: str = Field(min_length=1, max_length=256)


class CorrectionRequest(BaseModel):
    reason: str = Field(min_length=1, max_length=4000)
    corrected_revenue: dict[str, str | None]


def create_app(data_root: Path, service: WorkflowService | None = None) -> FastAPI:
    service = service or WorkflowService(data_root)
    tokens = SessionTokens(
        service.settings.session_secret_key,
        max_age_seconds=service.settings.session_max_age_seconds,
    )
    @asynccontextmanager
    async def lifespan(_: FastAPI):
        try:
            yield
        finally:
            service.close()

    app = FastAPI(title="仁恒合并利润表工作台", version="2.0.0", lifespan=lifespan)
    app.state.workflow_service = service
    web_dir = Path(__file__).with_name("web")
    app.mount("/static", StaticFiles(directory=web_dir), name="static")

    def current_user(request: Request) -> dict[str, Any]:
        user_id = tokens.read_user_id(request.cookies.get(SESSION_COOKIE))
        if not user_id:
            raise HTTPException(status_code=401, detail="请先登录")
        user = service.app_repository.get_user(user_id)
        if not user or not user.get("is_active"):
            raise HTTPException(status_code=401, detail="用户不存在或已停用")
        return user

    def role_guard(*allowed_roles: str) -> Callable[..., dict[str, Any]]:
        def dependency(user: dict[str, Any] = Depends(current_user)) -> dict[str, Any]:
            if user["role"] not in allowed_roles:
                raise HTTPException(status_code=403, detail="当前账号没有执行此操作的权限")
            return user
        return dependency

    @app.get("/", response_class=HTMLResponse)
    def index() -> HTMLResponse:
        return HTMLResponse((web_dir / "index.html").read_text(encoding="utf-8"))

    @app.post("/api/auth/login")
    def login(request: LoginRequest, response: Response) -> dict[str, Any]:
        user = service.authenticate(request.username, request.password)
        if not user:
            raise HTTPException(status_code=401, detail="用户名或密码错误")
        response.set_cookie(
            SESSION_COOKIE,
            tokens.issue(user),
            max_age=service.settings.session_max_age_seconds,
            httponly=True,
            secure=service.settings.session_cookie_secure,
            samesite="strict",
            path="/",
        )
        return {key: user[key] for key in ("username", "display_name", "role")}

    @app.post("/api/auth/logout")
    def logout(response: Response) -> dict[str, str]:
        response.delete_cookie(SESSION_COOKIE, path="/", httponly=True, samesite="strict")
        return {"status": "logged_out"}

    @app.get("/api/auth/me")
    def me(user: dict[str, Any] = Depends(current_user)) -> dict[str, Any]:
        return {key: user[key] for key in ("username", "display_name", "role")}

    @app.get("/api/runs")
    def list_runs(_: dict[str, Any] = Depends(current_user)) -> list[dict[str, Any]]:
        return service.list_runs()

    @app.post("/api/runs")
    async def create_run(
        year: int = Form(...),
        month: int = Form(...),
        files: list[UploadFile] = File(...),
        user: dict[str, Any] = Depends(role_guard("FINANCE", "REVIEWER", "ADMIN")),
    ) -> dict[str, Any]:
        try:
            return service.create_run(year=year, month=month, user=user, uploads=files)
        except (ValueError, OSError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error
        finally:
            for upload in files:
                await upload.close()

    @app.get("/api/runs/{run_id}")
    def get_run(run_id: str, _: dict[str, Any] = Depends(current_user)) -> dict[str, Any]:
        try:
            return service.get_status(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error

    @app.post("/api/runs/{run_id}/files")
    async def add_files(
        run_id: str,
        files: list[UploadFile] = File(...),
        user: dict[str, Any] = Depends(role_guard("FINANCE", "REVIEWER", "ADMIN")),
    ) -> dict[str, Any]:
        try:
            return service.add_uploads(run_id, files, user=user)
        except (KeyError, ValueError, OSError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error
        finally:
            for upload in files:
                await upload.close()

    @app.post("/api/runs/{run_id}/preflight")
    def preflight_run(
        run_id: str,
        _: dict[str, Any] = Depends(role_guard("FINANCE", "REVIEWER", "ADMIN")),
    ) -> dict[str, Any]:
        try:
            return service.preflight(run_id)
        except (KeyError, ValueError, RuntimeError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error

    @app.post("/api/runs/{run_id}/start")
    def start_run(
        run_id: str,
        _: dict[str, Any] = Depends(role_guard("FINANCE", "REVIEWER", "ADMIN")),
    ) -> dict[str, Any]:
        try:
            return service.start(run_id)
        except (KeyError, ValueError, RuntimeError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error

    @app.post("/api/runs/{run_id}/resume")
    def resume_run(
        run_id: str,
        request: CorrectionRequest,
        reviewer: dict[str, Any] = Depends(role_guard("REVIEWER", "ADMIN")),
    ) -> dict[str, Any]:
        try:
            return service.resume(run_id, request.model_dump(mode="json"), reviewer=reviewer)
        except (KeyError, ValueError, RuntimeError) as error:
            raise HTTPException(status_code=400, detail=str(error)) from error

    @app.get("/api/runs/{run_id}/artifacts/{key}")
    def download_artifact(
        run_id: str,
        key: str,
        _: dict[str, Any] = Depends(current_user),
    ) -> StreamingResponse:
        try:
            object_response, record = service.open_artifact(run_id, key)
        except (KeyError, FileNotFoundError) as error:
            raise HTTPException(status_code=404, detail=str(error)) from error

        def chunks():
            try:
                yield from object_response.stream(64 * 1024)
            finally:
                object_response.close()
                object_response.release_conn()

        filename = quote(record["original_filename"], safe="")
        return StreamingResponse(
            chunks(),
            media_type=record["content_type"],
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"},
        )

    return app
