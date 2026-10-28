"""协调PostgreSQL、MinIO、LangGraph与临时Excel工作区。"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
from threading import RLock
from typing import Any, Iterable
from urllib.parse import quote, unquote, urlsplit
from uuid import uuid4

from .database import RunRepository
from .security import hash_password, verify_password
from .settings import Settings
from .storage import MinioArtifactStore


class WorkflowService:
    """单机执行服务；同一进程串行运行Excel工作流。"""

    def __init__(
        self,
        data_root: Path,
        *,
        settings: Settings | None = None,
        repository: RunRepository | None = None,
        storage: Any | None = None,
        recover_pending: bool = True,
    ) -> None:
        self.data_root = data_root.resolve()
        self.runs_root = self.data_root / "runs"
        self.runs_root.mkdir(parents=True, exist_ok=True)
        self.settings = settings or Settings.from_env()
        self.app_repository = repository or RunRepository(self.settings.database_url)
        self.app_repository.validate_schema()
        self.storage = storage or MinioArtifactStore(
            self.settings.minio_endpoint,
            self.settings.minio_access_key,
            self.settings.minio_secret_key,
            self.settings.minio_bucket,
            secure=self.settings.minio_secure,
        )
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="profit-workflow")
        self._graph_lock = RLock()
        self._graph: Any | None = None
        self._checkpointer_context: Any | None = None
        if recover_pending:
            self.app_repository.fail_interrupted_uploads()
            self._enqueue_pending_runs()

    def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
        user = self.app_repository.get_user_by_username(username)
        if not user or not user.get("is_active") or not verify_password(password, user["password_hash"]):
            return None
        return {key: user[key] for key in ("user_id", "username", "display_name", "role", "is_active")}

    def _ensure_graph(self) -> Any:
        if self._graph is not None:
            return self._graph
        try:
            import os

            os.environ.setdefault("LANGGRAPH_STRICT_MSGPACK", "true")
            from langgraph.checkpoint.postgres import PostgresSaver
            from langgraph_profit_workflow import build_workflow
        except ImportError as error:
            raise RuntimeError(
                "网页服务需要PostgreSQL版LangGraph依赖，请执行 python -m pip install -r requirements.txt"
            ) from error
        self._checkpointer_context = PostgresSaver.from_conn_string(
            self.settings.checkpoint_database_url
        )
        checkpointer = self._checkpointer_context.__enter__()
        checkpointer.setup()
        self._graph = build_workflow(
            checkpointer,
            artifact_persister=self._persist_node_updates,
            artifact_resolver=self._resolve_artifact_reference,
        )
        return self._graph

    def _run_paths(self, run_id: str) -> tuple[Path, Path, Path]:
        run_dir = self.runs_root / run_id
        return run_dir, run_dir / "input", run_dir / "output"

    def _write_uploads(
        self,
        run_id: str,
        uploads: Iterable[Any],
        *,
        created_by: str,
        allow_overwrite: bool,
    ) -> list[dict[str, Any]]:
        run_dir, input_dir, _ = self._run_paths(run_id)
        input_dir.mkdir(parents=True, exist_ok=True)
        written: list[dict[str, Any]] = []
        seen: set[str] = set()
        for upload in uploads:
            filename = self._safe_filename(getattr(upload, "filename", ""))
            normalized_name = filename.casefold()
            if normalized_name in seen:
                raise ValueError(f"上传文件重名：{filename}")
            seen.add(normalized_name)
            relative_path = (Path("input") / filename).as_posix()
            target = run_dir / relative_path
            if target.exists() and not allow_overwrite:
                raise ValueError(f"上传文件重名：{filename}")
            file_id = str(uuid4())
            object_key = f"runs/{run_id}/inputs/{file_id}/{quote(filename, safe='._-()') }"
            metadata = self.storage.put_upload(upload, object_key)
            try:
                record = self.app_repository.add_file(
                    run_id=run_id,
                    file_id=file_id,
                    role="INPUT",
                    artifact_key=f"input:{filename.casefold()}",
                    object_key=object_key,
                    original_filename=filename,
                    local_path=relative_path,
                    content_type=metadata["content_type"],
                    size_bytes=metadata["size_bytes"],
                    sha256=metadata["sha256"],
                    created_by=created_by,
                )
            except Exception:
                try:
                    self.storage.delete(object_key)
                except Exception:
                    pass
                raise
            self.storage.download_to_path(object_key, target)
            written.append(record)
        return written

    @staticmethod
    def _safe_filename(filename: str) -> str:
        name = str(filename or "upload.bin").replace("\\", "/").split("/")[-1]
        name = re.sub(r"[^0-9A-Za-z一-龥._()（） -]", "_", name).strip(" .")
        if not name or name in {".", ".."}:
            raise ValueError("上传文件名无效")
        return name

    def create_run(
        self,
        *,
        year: int,
        month: int,
        user: dict[str, Any],
        uploads: Iterable[Any],
    ) -> dict[str, Any]:
        if year < 1900 or year > 2200:
            raise ValueError("年度超出有效范围")
        if not 1 <= month <= 12:
            raise ValueError("月份必须为1至12")
        run_id = f"RH-{year:04d}-{month:02d}-{uuid4().hex[:8].upper()}"
        thread_id = uuid4().hex
        self.app_repository.create_run(
            run_id=run_id,
            thread_id=thread_id,
            year=year,
            month=month,
            created_by=user["user_id"],
        )
        try:
            records = self._write_uploads(
                run_id, uploads, created_by=user["user_id"], allow_overwrite=False
            )
            if not records:
                raise ValueError("至少上传一个资料文件")
            self.app_repository.finish_uploads(run_id)
        except Exception as error:
            self.app_repository.update(
                run_id, status="UPLOAD_FAILED", error_message=f"{type(error).__name__}: {error}"
            )
            raise ValueError(f"上传失败，任务{run_id}已标记为UPLOAD_FAILED：{error}") from error
        result = self._require_run(run_id)
        result["uploaded"] = [record["original_filename"] for record in records]
        result["files"] = [_public_file(record) for record in records]
        return result

    def add_uploads(
        self, run_id: str, uploads: Iterable[Any], *, user: dict[str, Any]
    ) -> dict[str, Any]:
        row = self._require_run(run_id)
        if row["status"] not in {"DRAFT", "UPLOAD_FAILED"}:
            raise ValueError("只有DRAFT或UPLOAD_FAILED任务可以补充或替换资料")
        self._hydrate_files(run_id)
        self.app_repository.update(
            run_id, preflight_status="NOT_RUN", preflight_result=None, error_message=None
        )
        try:
            records = self._write_uploads(
                run_id, uploads, created_by=user["user_id"], allow_overwrite=True
            )
        except Exception as error:
            self.app_repository.update(run_id, error_message=f"{type(error).__name__}: {error}")
            raise
        if not records:
            raise ValueError("至少上传一个资料文件")
        self.app_repository.finish_uploads(run_id)
        result = self._require_run(run_id)
        result["uploaded"] = [record["original_filename"] for record in records]
        return result

    def _preflight_state(self, row: dict[str, Any]) -> dict[str, Any]:
        from langgraph_profit_workflow import determine_period_policy, preflight

        _, source_dir, output_root = self._run_paths(row["run_id"])
        base = {
            "year": row["year"],
            "month": row["month"],
            "source_dir": str(source_dir),
            "output_root": str(output_root),
        }
        policy_state = {**base, **determine_period_policy(base)}
        return {**policy_state, **preflight(policy_state)}

    def preflight(self, run_id: str) -> dict[str, Any]:
        row = self._require_run(run_id)
        if row["status"] not in {"DRAFT", "FAILED"}:
            raise ValueError(f"当前状态不能预检：{row['status']}")
        self._hydrate_files(run_id)
        self.app_repository.update(run_id, status="VALIDATING", error_message=None)
        try:
            state = self._preflight_state(row)
            selected = state["selected_files"]
            payload = {
                "ok": True,
                "files": [
                    {"role": role, "filename": Path(path).name}
                    for role, value in selected.items()
                    for path in (value if isinstance(value, list) else [value])
                ],
                "selected_files": {
                    role: [Path(path).name for path in value]
                    if isinstance(value, list)
                    else Path(value).name
                    for role, value in selected.items()
                },
                "current_input_fingerprint": state["current_input_fingerprint"],
                "prior_input_fingerprint": state["prior_input_fingerprint"],
            }
            run_dir, _, _ = self._run_paths(run_id)
            preflight_path = run_dir / "preflight.json"
            preflight_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            self._persist_local_file(
                run_id,
                artifact_key="preflight",
                path=preflight_path,
                created_by=row["created_by"],
                role="SYSTEM",
            )
            self.app_repository.update(
                run_id,
                status="DRAFT",
                preflight_status="PASSED",
                preflight_result=payload,
                error_message=None,
            )
        except Exception as error:
            payload = {"ok": False, "error": f"{type(error).__name__}: {error}"}
            self.app_repository.update(
                run_id,
                status="DRAFT",
                preflight_status="FAILED",
                preflight_result=payload,
                error_message=payload["error"],
            )
        return self.get_status(run_id)

    def start(self, run_id: str) -> dict[str, Any]:
        row = self._require_run(run_id)
        if row["status"] not in {"DRAFT", "FAILED"}:
            raise ValueError(f"当前状态不能启动：{row['status']}")
        if row.get("preflight_status") != "PASSED":
            row = self.preflight(run_id)
            if row.get("preflight_status") != "PASSED":
                raise ValueError(row.get("error_message") or "资料预检未通过")
        self._hydrate_files(run_id)
        self.app_repository.queue_start(run_id)
        self._executor.submit(self._invoke_start, self._require_run(run_id))
        return self.get_status(run_id)

    def resume(
        self, run_id: str, correction: dict[str, Any], *, reviewer: dict[str, Any]
    ) -> dict[str, Any]:
        row = self._require_run(run_id)
        if row["status"] != "WAITING_REVIEW":
            raise ValueError(f"当前状态不能恢复：{row['status']}")
        state = self._state(row)
        review_round = int(state.get("review_round", 0))
        reason = str(correction.get("reason", "")).strip()
        corrected = correction.get("corrected_revenue")
        if not reason or not isinstance(corrected, dict):
            raise ValueError("必须填写审核原因和营业收入修正值")
        self.app_repository.create_approval_and_queue(
            run_id=run_id,
            review_round=review_round,
            reviewer_id=reviewer["user_id"],
            reason=reason,
            corrected_revenue=corrected,
            before_differences=state.get("m7_differences", []),
        )
        correction = {**correction, "reviewed_by": reviewer["display_name"]}
        self._executor.submit(self._invoke_resume, self._require_run(run_id), correction)
        return self.get_status(run_id)

    def _config(self, row: dict[str, Any]) -> dict[str, Any]:
        return {"configurable": {"thread_id": row["thread_id"], "run_id": row["run_id"]}}

    def _workflow_input(self, row: dict[str, Any]) -> dict[str, Any]:
        _, source_dir, output_root = self._run_paths(row["run_id"])
        return {
            "year": row["year"],
            "month": row["month"],
            "source_dir": str(source_dir),
            "output_root": str(output_root),
        }

    def _invoke_start(self, row: dict[str, Any]) -> None:
        self.app_repository.update(row["run_id"], status="RUNNING")
        try:
            self._hydrate_files(row["run_id"])
            with self._graph_lock:
                result = self._ensure_graph().invoke(
                    self._workflow_input(row), config=self._config(row)
                )
            self._sync_result(row, result)
        except Exception as error:
            self._record_failure(row, error)

    def _invoke_resume(self, row: dict[str, Any], correction: dict[str, Any]) -> None:
        self.app_repository.update(row["run_id"], status="RUNNING")
        try:
            from langgraph.types import Command

            self._hydrate_files(row["run_id"])
            with self._graph_lock:
                result = self._ensure_graph().invoke(
                    Command(resume=correction), config=self._config(row)
                )
            self._sync_result(row, result)
        except Exception as error:
            self._record_failure(row, error)

    def _persist_node_updates(self, run_id: str, updates: dict[str, Any]) -> dict[str, Any]:
        updates = dict(updates or {})
        path_to_ref: dict[str, str] = {}
        artifact_updates = dict(updates.get("artifacts", {}))
        row = self._require_run(run_id)
        for key, value in artifact_updates.items():
            path = Path(value)
            role = "FORMAL_REPORT" if key == "formal_report" else "REVIEW" if key.startswith("review_round_") else "ARTIFACT"
            record = self._persist_local_file(
                run_id,
                artifact_key=key,
                path=path,
                created_by=row["created_by"],
                role=role,
            )
            reference = _object_ref(self.settings.minio_bucket, record["object_key"])
            path_to_ref[str(path.resolve())] = reference
            artifact_updates[key] = reference
        updates["artifacts"] = artifact_updates
        for field, fallback_key in (("difference_review_path", "difference_review"), ("formal_report", "formal_report")):
            value = updates.get(field)
            if value and not str(value).startswith("minio://"):
                reference = path_to_ref.get(str(Path(value).resolve()))
                if reference is None:
                    record = self._persist_local_file(
                        run_id,
                        artifact_key=fallback_key,
                        path=Path(value),
                        created_by=row["created_by"],
                        role="FORMAL_REPORT" if field == "formal_report" else "REVIEW",
                    )
                    reference = _object_ref(self.settings.minio_bucket, record["object_key"])
                updates[field] = reference
        return updates

    def _persist_local_file(
        self,
        run_id: str,
        *,
        artifact_key: str,
        path: Path,
        created_by: str,
        role: str,
    ) -> dict[str, Any]:
        run_dir, _, _ = self._run_paths(run_id)
        resolved = path.resolve()
        try:
            relative = resolved.relative_to(run_dir.resolve())
        except ValueError as error:
            raise ValueError("工作流产物路径超出当前任务目录") from error
        current = self.app_repository.get_artifact(run_id, artifact_key)
        if current and current["sha256"] == _sha256_file(resolved):
            return current
        file_id = str(uuid4())
        object_key = f"runs/{run_id}/artifacts/{quote(artifact_key, safe='._-')}/{file_id}/{quote(resolved.name, safe='._-()') }"
        metadata = self.storage.put_path(resolved, object_key, _content_type(resolved))
        try:
            return self.app_repository.add_file(
                run_id=run_id,
                file_id=file_id,
                role=role,
                artifact_key=artifact_key,
                object_key=object_key,
                original_filename=resolved.name,
                local_path=relative.as_posix(),
                content_type=metadata["content_type"],
                size_bytes=metadata["size_bytes"],
                sha256=metadata["sha256"],
                created_by=created_by,
            )
        except Exception:
            try:
                self.storage.delete(object_key)
            except Exception:
                pass
            raise

    def _resolve_artifact_reference(self, run_id: str, reference: str) -> Path:
        parsed = _parse_object_ref(reference)
        if parsed is None:
            return Path(reference)
        bucket, object_key = parsed
        if bucket != self.settings.minio_bucket:
            raise ValueError("State中的文件引用不属于当前MinIO bucket")
        record = self.app_repository.get_file_by_object_key(run_id, object_key)
        if not record:
            raise FileNotFoundError("Checkpoint引用的MinIO文件索引不存在")
        return self._materialize_file(run_id, record)

    def _materialize_file(self, run_id: str, record: dict[str, Any]) -> Path:
        run_dir, _, _ = self._run_paths(run_id)
        target = (run_dir / record["local_path"]).resolve()
        if run_dir.resolve() not in target.parents:
            raise ValueError("数据库中的本地工作路径无效")
        if not target.is_file() or target.stat().st_size != record["size_bytes"] or _sha256_file(target) != record["sha256"]:
            self.storage.download_to_path(record["object_key"], target)
        return target

    def _hydrate_files(self, run_id: str) -> None:
        for record in self.app_repository.list_files(run_id, current_only=True):
            self._materialize_file(run_id, record)

    def _sync_result(self, row: dict[str, Any], result: Any) -> None:
        state = self._state(row)
        if isinstance(result, dict) and result.get("__interrupt__"):
            status = "WAITING_REVIEW"
        else:
            status = state.get("workflow_status", result.get("workflow_status", "RUNNING"))
        updates: dict[str, Any] = {"status": status}
        formal_ref = state.get("formal_report")
        if formal_ref:
            parsed = _parse_object_ref(formal_ref)
            if parsed:
                updates["formal_report_object_key"] = parsed[1]
        if status in {"SUCCEEDED", "FAILED"}:
            updates["completed_at"] = datetime.now(timezone.utc)
        self.app_repository.update(row["run_id"], **updates)

    def _record_failure(self, row: dict[str, Any], error: Exception) -> None:
        self.app_repository.update(
            row["run_id"],
            status="FAILED",
            error_message=f"{type(error).__name__}: {error}",
            completed_at=datetime.now(timezone.utc),
        )

    def _state(self, row: dict[str, Any]) -> dict[str, Any]:
        try:
            with self._graph_lock:
                snapshot = self._ensure_graph().get_state(self._config(row))
            return dict(snapshot.values)
        except Exception:
            return {}

    def get_status(self, run_id: str) -> dict[str, Any]:
        row = self._require_run(run_id)
        state = self._state(row) if row["status"] in {"WAITING_REVIEW", "SUCCEEDED", "FAILED"} else {}
        preflight = row.get("preflight_result")
        artifacts = {
            record["artifact_key"]: record["original_filename"]
            for record in self.app_repository.list_files(run_id, current_only=True)
            if record["role"] in {"ARTIFACT", "REVIEW", "FORMAL_REPORT"}
        }
        row.update(
            {
                "preflight": preflight,
                "workflow_status": state.get("workflow_status", row["status"]),
                "review_round": state.get("review_round", 0),
                "m7_differences": state.get("m7_differences", []),
                "subject_differences": state.get("subject_differences", []),
                "unbilled_evidence": state.get("unbilled_evidence", []),
                "difference_review_path": None,
                "artifacts": artifacts,
                "formal_report": artifacts.get("formal_report"),
                "error": state.get("error"),
                "approvals": self.app_repository.list_approvals(run_id),
                "files": [
                    _public_file(record)
                    for record in self.app_repository.list_files(run_id, current_only=True)
                ],
            }
        )
        row.pop("formal_report_object_key", None)
        return row

    def list_runs(self) -> list[dict[str, Any]]:
        rows = self.app_repository.list()
        for row in rows:
            row.pop("formal_report_object_key", None)
        return rows

    def artifact_info(self, run_id: str, key: str) -> dict[str, Any]:
        self._require_run(run_id)
        record = self.app_repository.get_artifact(run_id, key)
        if not record:
            raise FileNotFoundError(f"产物不存在：{key}")
        return record

    def open_artifact(self, run_id: str, key: str) -> tuple[Any, dict[str, Any]]:
        record = self.artifact_info(run_id, key)
        return self.storage.open_download(record["object_key"]), record

    def _require_run(self, run_id: str) -> dict[str, Any]:
        row = self.app_repository.get(run_id)
        if row is None:
            raise KeyError(f"运行不存在：{run_id}")
        return row

    def _enqueue_pending_runs(self) -> None:
        for row in self.app_repository.list_incomplete():
            self._executor.submit(self._recover_run, row)

    def _recover_run(self, row: dict[str, Any]) -> None:
        try:
            self._hydrate_files(row["run_id"])
            graph = self._ensure_graph()
            with self._graph_lock:
                snapshot = graph.get_state(self._config(row))
                state = dict(snapshot.values)
                if state.get("workflow_status") == "WAITING_REVIEW":
                    self.app_repository.update(row["run_id"], status="WAITING_REVIEW")
                    return
                if state.get("workflow_status") in {"SUCCEEDED", "FAILED"}:
                    self._sync_result(row, state)
                    return
                if row["status"] == "QUEUED" and snapshot.next and state:
                    approval = self.app_repository.latest_approval(row["run_id"])
                    if approval:
                        from langgraph.types import Command
                        correction = {
                            "reviewed_by": approval["reviewer"],
                            "reason": approval["reason"],
                            "corrected_revenue": approval["corrected_revenue"],
                        }
                        result = graph.invoke(Command(resume=correction), config=self._config(row))
                    else:
                        result = graph.invoke(None, config=self._config(row))
                elif state and snapshot.next:
                    result = graph.invoke(None, config=self._config(row))
                else:
                    result = graph.invoke(self._workflow_input(row), config=self._config(row))
            self._sync_result(row, result)
        except Exception as error:
            self._record_failure(row, error)

    def close(self) -> None:
        self._executor.shutdown(wait=True)
        if self._checkpointer_context is not None:
            self._checkpointer_context.__exit__(None, None, None)
        self.app_repository.close()


def _object_ref(bucket: str, object_key: str) -> str:
    return f"minio://{bucket}/{quote(object_key, safe='/') }"


def _parse_object_ref(reference: str) -> tuple[str, str] | None:
    if not reference.startswith("minio://"):
        return None
    parsed = urlsplit(reference)
    return parsed.netloc, unquote(parsed.path.lstrip("/"))


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while True:
            chunk = stream.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _content_type(path: Path) -> str:
    suffix = path.suffix.lower()
    return {
        ".xls": "application/vnd.ms-excel",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".json": "application/json",
        ".txt": "text/plain; charset=utf-8",
    }.get(suffix, "application/octet-stream")


def _public_file(record: dict[str, Any]) -> dict[str, Any]:
    return {
        key: record[key]
        for key in (
            "file_id", "role", "artifact_key", "original_filename", "content_type",
            "size_bytes", "sha256", "version", "is_current", "created_at",
        )
        if key in record
    }
