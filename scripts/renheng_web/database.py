"""单机网页层的运行记录数据库。

LangGraph 自己管理 checkpoint 表；本模块只保存页面需要的运行索引，
不直接依赖 LangGraph 的内部表结构。
"""
from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any


class RunRepository:
    def __init__(self, database_path: Path) -> None:
        self.database_path = database_path
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path, timeout=30)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS workflow_runs (
                    run_id TEXT PRIMARY KEY,
                    thread_id TEXT NOT NULL UNIQUE,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    operator TEXT NOT NULL,
                    status TEXT NOT NULL,
                    source_dir TEXT NOT NULL,
                    output_root TEXT NOT NULL,
                    run_output_dir TEXT,
                    formal_report TEXT,
                    error_message TEXT,
                    preflight_status TEXT NOT NULL DEFAULT 'NOT_RUN',
                    preflight_result TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    completed_at TEXT
                )
                """
            )
            # 兼容网页层首次升级前已经创建的 app.sqlite。
            columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(workflow_runs)").fetchall()
            }
            if "preflight_status" not in columns:
                connection.execute(
                    "ALTER TABLE workflow_runs ADD COLUMN preflight_status TEXT NOT NULL DEFAULT 'NOT_RUN'"
                )
            if "preflight_result" not in columns:
                connection.execute(
                    "ALTER TABLE workflow_runs ADD COLUMN preflight_result TEXT"
                )

    def create_run(
        self,
        *,
        run_id: str,
        thread_id: str,
        year: int,
        month: int,
        operator: str,
        source_dir: Path,
        output_root: Path,
    ) -> dict[str, Any]:
        now = datetime.now().isoformat(timespec="seconds")
        values = {
            "run_id": run_id,
            "thread_id": thread_id,
            "year": year,
            "month": month,
            "operator": operator,
            "status": "DRAFT",
            "source_dir": str(source_dir),
            "output_root": str(output_root),
            "run_output_dir": None,
            "formal_report": None,
            "error_message": None,
            "preflight_status": "NOT_RUN",
            "preflight_result": None,
            "created_at": now,
            "updated_at": now,
            "completed_at": None,
        }
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO workflow_runs
                (run_id, thread_id, year, month, operator, status, source_dir,
                 output_root, run_output_dir, formal_report, error_message,
                 preflight_status, preflight_result,
                 created_at, updated_at, completed_at)
                VALUES
                (:run_id, :thread_id, :year, :month, :operator, :status, :source_dir,
                 :output_root, :run_output_dir, :formal_report, :error_message,
                 :preflight_status, :preflight_result,
                 :created_at, :updated_at, :completed_at)
                """,
                values,
            )
        return values

    def get(self, run_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM workflow_runs WHERE run_id = ?", (run_id,)
            ).fetchone()
        return None if row is None else dict(row)

    def list(self, limit: int = 50) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM workflow_runs ORDER BY created_at DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(row) for row in rows]

    def update(self, run_id: str, **fields: Any) -> dict[str, Any] | None:
        allowed = {
            "status", "run_output_dir", "formal_report", "error_message", "completed_at",
            "preflight_status", "preflight_result",
        }
        updates = {key: value for key, value in fields.items() if key in allowed}
        if not updates:
            return self.get(run_id)
        updates["updated_at"] = datetime.now().isoformat(timespec="seconds")
        assignments = ", ".join(f"{key} = :{key}" for key in updates)
        updates["run_id"] = run_id
        with self._connect() as connection:
            connection.execute(
                f"UPDATE workflow_runs SET {assignments} WHERE run_id = :run_id", updates
            )
        return self.get(run_id)
