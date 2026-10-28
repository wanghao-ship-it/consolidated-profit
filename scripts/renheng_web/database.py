"""PostgreSQL业务仓储；不直接访问LangGraph管理的Checkpoint表。"""
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from uuid import uuid4

from sqlalchemy import Engine, create_engine, inspect, select, update
from sqlalchemy.pool import StaticPool

from .db_schema import approval_records, metadata, run_events, users, workflow_files, workflow_runs


class RunRepository:
    def __init__(
        self,
        database_url: str | Path | Engine,
        *,
        create_schema: bool = False,
    ) -> None:
        if isinstance(database_url, Engine):
            self.engine = database_url
            self._owns_engine = False
        else:
            url = str(database_url)
            options: dict[str, Any] = {"pool_pre_ping": True}
            if url in {"sqlite://", "sqlite:///:memory:"}:
                options.update(connect_args={"check_same_thread": False}, poolclass=StaticPool)
            self.engine = create_engine(url, **options)
            self._owns_engine = True
        if create_schema:
            metadata.create_all(self.engine)

    def validate_schema(self) -> None:
        required = {"users", "workflow_runs", "workflow_files", "approval_records", "run_events"}
        with self.engine.connect() as connection:
            existing = set(inspect(connection).get_table_names())
        missing = required - existing
        if missing:
            raise RuntimeError(
                "PostgreSQL业务表尚未初始化；请先执行 python scripts/run_db_migrations.py。"
            )

    def create_user(
        self,
        *,
        username: str,
        display_name: str,
        role: str,
        password_hash: str,
    ) -> dict[str, Any]:
        now = _now()
        value = {
            "user_id": str(uuid4()),
            "username": username.strip(),
            "display_name": display_name.strip(),
            "role": role,
            "password_hash": password_hash,
            "is_active": True,
            "created_at": now,
        }
        with self.engine.begin() as connection:
            connection.execute(users.insert().values(**value))
        return _user_public(value)

    def get_user_by_username(self, username: str) -> dict[str, Any] | None:
        with self.engine.connect() as connection:
            row = connection.execute(
                select(users).where(users.c.username == username.strip())
            ).mappings().first()
        return None if row is None else dict(row)

    def get_user(self, user_id: str) -> dict[str, Any] | None:
        with self.engine.connect() as connection:
            row = connection.execute(select(users).where(users.c.user_id == user_id)).mappings().first()
        return None if row is None else _user_public(dict(row))

    def create_run(
        self,
        *,
        run_id: str,
        thread_id: str,
        year: int,
        month: int,
        created_by: str,
    ) -> dict[str, Any]:
        now = _now()
        value = {
            "run_id": run_id,
            "thread_id": thread_id,
            "year": year,
            "month": month,
            "created_by": created_by,
            "status": "UPLOADING",
            "preflight_status": "NOT_RUN",
            "preflight_result": None,
            "formal_report_object_key": None,
            "error_message": None,
            "created_at": now,
            "updated_at": now,
            "completed_at": None,
        }
        with self.engine.begin() as connection:
            connection.execute(workflow_runs.insert().values(**value))
            self._event(connection, run_id, created_by, "RUN_CREATED", {"status": "UPLOADING"})
        return self.get(run_id) or value

    def add_file(
        self,
        *,
        run_id: str,
        file_id: str,
        role: str,
        artifact_key: str,
        object_key: str,
        original_filename: str,
        local_path: str,
        content_type: str,
        size_bytes: int,
        sha256: str,
        created_by: str,
    ) -> dict[str, Any]:
        now = _now()
        with self.engine.begin() as connection:
            current = connection.execute(
                select(workflow_files)
                .where(
                    workflow_files.c.run_id == run_id,
                    workflow_files.c.role == role,
                    workflow_files.c.artifact_key == artifact_key,
                    workflow_files.c.is_current.is_(True),
                )
                .with_for_update()
            ).mappings().first()
            version = int(current["version"]) + 1 if current else 1
            if current:
                connection.execute(
                    update(workflow_files)
                    .where(workflow_files.c.file_id == current["file_id"])
                    .values(is_current=False)
                )
            value = {
                "file_id": file_id,
                "run_id": run_id,
                "role": role,
                "artifact_key": artifact_key,
                "object_key": object_key,
                "original_filename": original_filename,
                "local_path": local_path,
                "content_type": content_type,
                "size_bytes": size_bytes,
                "sha256": sha256,
                "version": version,
                "is_current": True,
                "created_by": created_by,
                "created_at": now,
            }
            connection.execute(workflow_files.insert().values(**value))
            self._event(
                connection,
                run_id,
                created_by,
                "FILE_STORED",
                {"file_id": file_id, "role": role, "artifact_key": artifact_key, "version": version},
            )
        return _json_row(value)

    def finish_uploads(self, run_id: str) -> dict[str, Any] | None:
        return self.update(run_id, status="DRAFT", error_message=None)

    def fail_interrupted_uploads(self) -> int:
        """标记上次进程退出时遗留的上传任务，避免任务永久停在UPLOADING。"""
        now = _now()
        with self.engine.begin() as connection:
            rows = connection.execute(
                select(workflow_runs.c.run_id).where(workflow_runs.c.status == "UPLOADING")
            ).all()
            if not rows:
                return 0
            run_ids = [row.run_id for row in rows]
            connection.execute(
                update(workflow_runs)
                .where(workflow_runs.c.run_id.in_(run_ids))
                .values(
                    status="UPLOAD_FAILED",
                    error_message="服务在上传完成前退出；请检查资料并重新上传后继续。",
                    updated_at=now,
                )
            )
            for run_id in run_ids:
                self._event(
                    connection,
                    run_id,
                    None,
                    "UPLOAD_INTERRUPTED",
                    {"status": "UPLOAD_FAILED"},
                )
        return len(run_ids)

    def queue_start(self, run_id: str) -> None:
        now = _now()
        with self.engine.begin() as connection:
            result = connection.execute(
                update(workflow_runs)
                .where(
                    workflow_runs.c.run_id == run_id,
                    workflow_runs.c.status.in_(["DRAFT", "FAILED"]),
                    workflow_runs.c.preflight_status == "PASSED",
                )
                .values(status="QUEUED", error_message=None, updated_at=now)
            )
            if result.rowcount != 1:
                raise ValueError("任务状态已变化或预检未通过，不能启动")
            self._event(connection, run_id, None, "RUN_QUEUED", {})

    def get(self, run_id: str) -> dict[str, Any] | None:
        with self.engine.connect() as connection:
            row = connection.execute(
                select(
                    workflow_runs,
                    users.c.display_name.label("operator"),
                    users.c.username.label("operator_username"),
                )
                .select_from(workflow_runs.join(users, workflow_runs.c.created_by == users.c.user_id))
                .where(workflow_runs.c.run_id == run_id)
            ).mappings().first()
        return None if row is None else _json_row(dict(row))

    def list(self, limit: int = 50) -> list[dict[str, Any]]:
        with self.engine.connect() as connection:
            rows = connection.execute(
                select(
                    workflow_runs,
                    users.c.display_name.label("operator"),
                    users.c.username.label("operator_username"),
                )
                .select_from(workflow_runs.join(users, workflow_runs.c.created_by == users.c.user_id))
                .order_by(workflow_runs.c.created_at.desc())
                .limit(max(1, min(limit, 200)))
            ).mappings().all()
        return [_json_row(dict(row)) for row in rows]

    def update(self, run_id: str, **fields: Any) -> dict[str, Any] | None:
        allowed = {
            "status", "formal_report_object_key", "error_message",
            "completed_at", "preflight_status", "preflight_result",
        }
        updates = {key: value for key, value in fields.items() if key in allowed}
        if not updates:
            return self.get(run_id)
        updates["updated_at"] = _now()
        with self.engine.begin() as connection:
            current = connection.execute(
                select(workflow_runs.c.status, workflow_runs.c.created_by)
                .where(workflow_runs.c.run_id == run_id)
                .with_for_update()
            ).mappings().first()
            if current is None:
                return None
            connection.execute(update(workflow_runs).where(workflow_runs.c.run_id == run_id).values(**updates))
            if "status" in updates and updates["status"] != current["status"]:
                self._event(
                    connection,
                    run_id,
                    None,
                    "STATUS_CHANGED",
                    {"from": current["status"], "to": updates["status"]},
                )
        return self.get(run_id)

    def create_approval_and_queue(
        self,
        *,
        run_id: str,
        review_round: int,
        reviewer_id: str,
        reason: str,
        corrected_revenue: dict[str, Any],
        before_differences: list[dict[str, Any]],
    ) -> dict[str, Any]:
        now = _now()
        approval_id = str(uuid4())
        with self.engine.begin() as connection:
            result = connection.execute(
                update(workflow_runs)
                .where(workflow_runs.c.run_id == run_id, workflow_runs.c.status == "WAITING_REVIEW")
                .values(status="QUEUED", error_message=None, updated_at=now)
            )
            if result.rowcount != 1:
                raise ValueError("任务状态已变化或不在等待审核状态，不能重复提交")
            connection.execute(
                approval_records.insert().values(
                    approval_id=approval_id,
                    run_id=run_id,
                    review_round=review_round,
                    reviewer_id=reviewer_id,
                    reason=reason,
                    corrected_revenue=corrected_revenue,
                    before_differences=before_differences,
                    created_at=now,
                )
            )
            self._event(
                connection,
                run_id,
                reviewer_id,
                "CORRECTION_SUBMITTED",
                {"approval_id": approval_id, "review_round": review_round},
            )
        return {"approval_id": approval_id, "run_id": run_id, "review_round": review_round}

    def list_files(self, run_id: str, *, current_only: bool = True) -> list[dict[str, Any]]:
        query = select(workflow_files).where(workflow_files.c.run_id == run_id)
        if current_only:
            query = query.where(workflow_files.c.is_current.is_(True))
        query = query.order_by(workflow_files.c.created_at, workflow_files.c.file_id)
        with self.engine.connect() as connection:
            rows = connection.execute(query).mappings().all()
        return [_json_row(dict(row)) for row in rows]

    def get_file_by_object_key(self, run_id: str, object_key: str) -> dict[str, Any] | None:
        with self.engine.connect() as connection:
            row = connection.execute(
                select(workflow_files).where(
                    workflow_files.c.run_id == run_id,
                    workflow_files.c.object_key == object_key,
                )
            ).mappings().first()
        return None if row is None else _json_row(dict(row))

    def get_artifact(self, run_id: str, artifact_key: str) -> dict[str, Any] | None:
        with self.engine.connect() as connection:
            row = connection.execute(
                select(workflow_files).where(
                    workflow_files.c.run_id == run_id,
                    workflow_files.c.role.in_(["ARTIFACT", "REVIEW", "FORMAL_REPORT", "SYSTEM"]),
                    workflow_files.c.artifact_key == artifact_key,
                    workflow_files.c.is_current.is_(True),
                )
            ).mappings().first()
        return None if row is None else _json_row(dict(row))

    def list_approvals(self, run_id: str) -> list[dict[str, Any]]:
        with self.engine.connect() as connection:
            rows = connection.execute(
                select(
                    approval_records,
                    users.c.display_name.label("reviewer"),
                    users.c.username.label("reviewer_username"),
                )
                .select_from(approval_records.join(users, approval_records.c.reviewer_id == users.c.user_id))
                .where(approval_records.c.run_id == run_id)
                .order_by(approval_records.c.review_round)
            ).mappings().all()
        return [_json_row(dict(row)) for row in rows]

    def latest_approval(self, run_id: str) -> dict[str, Any] | None:
        with self.engine.connect() as connection:
            row = connection.execute(
                select(
                    approval_records,
                    users.c.display_name.label("reviewer"),
                )
                .select_from(approval_records.join(users, approval_records.c.reviewer_id == users.c.user_id))
                .where(approval_records.c.run_id == run_id)
                .order_by(approval_records.c.review_round.desc())
                .limit(1)
            ).mappings().first()
        return None if row is None else _json_row(dict(row))

    def list_incomplete(self) -> list[dict[str, Any]]:
        with self.engine.connect() as connection:
            rows = connection.execute(
                select(
                    workflow_runs,
                    users.c.display_name.label("operator"),
                    users.c.username.label("operator_username"),
                )
                .select_from(workflow_runs.join(users, workflow_runs.c.created_by == users.c.user_id))
                .where(workflow_runs.c.status.in_(["QUEUED", "RUNNING"]))
                .order_by(workflow_runs.c.created_at)
            ).mappings().all()
        return [_json_row(dict(row)) for row in rows]

    def _event(
        self,
        connection: Any,
        run_id: str,
        actor_id: str | None,
        event_type: str,
        details: dict[str, Any],
    ) -> None:
        connection.execute(
            run_events.insert().values(
                event_id=str(uuid4()),
                run_id=run_id,
                actor_id=actor_id,
                event_type=event_type,
                details=details,
                created_at=_now(),
            )
        )

    def close(self) -> None:
        if self._owns_engine:
            self.engine.dispose()


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _json_row(value: dict[str, Any]) -> dict[str, Any]:
    return {key: item.isoformat() if isinstance(item, datetime) else item for key, item in value.items()}


def _user_public(value: dict[str, Any]) -> dict[str, Any]:
    return _json_row({key: item for key, item in value.items() if key != "password_hash"})
