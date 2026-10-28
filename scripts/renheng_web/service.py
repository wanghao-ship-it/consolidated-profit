"""把网页请求转换为 LangGraph 启动、查询和恢复操作。"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import json
from pathlib import Path
import re
import shutil
import sqlite3
from threading import RLock
from typing import Any, Iterable
from uuid import uuid4

from .database import RunRepository


class WorkflowService:
    """单机工作流服务。

    一个进程内只维护一个 LangGraph/SQLite Checkpointer，并用单线程执行器
    串行运行财务任务，避免多个 Excel 任务同时写入同一台机器。
    """

    def __init__(self, data_root: Path) -> None:
        self.data_root = data_root.resolve()
        self.runs_root = self.data_root / "runs"
        self.runs_root.mkdir(parents=True, exist_ok=True)
        self.app_repository = RunRepository(self.data_root / "app.sqlite")
        self.checkpoint_path = self.data_root / "langgraph_checkpoints.sqlite"
        # 先创建文件，真正的 checkpoints/writes 表由 LangGraph 首次实例化时建表。
        with sqlite3.connect(self.checkpoint_path):
            pass
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="profit-workflow")
        self._graph_lock = RLock()
        self._connection: sqlite3.Connection | None = None
        self._graph: Any | None = None

    def _ensure_graph(self) -> Any:
        if self._graph is not None:
            return self._graph
        try:
            from langgraph.checkpoint.sqlite import SqliteSaver
            from langgraph_profit_workflow import build_workflow
        except ImportError as error:
            raise RuntimeError(
                "网页服务需要 LangGraph 依赖，请先执行 python -m pip install -r requirements.txt"
            ) from error
        self._connection = sqlite3.connect(self.checkpoint_path, check_same_thread=False)
        self._graph = build_workflow(SqliteSaver(self._connection))
        return self._graph

    def _run_paths(self, run_id: str) -> tuple[Path, Path]:
        run_dir = self.runs_root / run_id
        return run_dir, run_dir / "input"

    def _write_uploads(
        self, input_dir: Path, uploads: Iterable[Any], *, allow_overwrite: bool
    ) -> list[str]:
        uploaded: list[str] = []
        for upload in uploads:
            filename = self._safe_filename(getattr(upload, "filename", ""))
            if filename in uploaded:
                raise ValueError(f"上传文件重名：{filename}")
            target = input_dir / filename
            if target.exists() and not allow_overwrite:
                raise ValueError(f"上传文件重名：{filename}")
            content = upload.file.read()
            if not content:
                raise ValueError(f"上传文件为空：{filename}")
            target.write_bytes(content)
            uploaded.append(filename)
        return uploaded

    @staticmethod
    def _safe_filename(filename: str) -> str:
        name = Path(filename or "upload.bin").name
        name = re.sub(r"[^0-9A-Za-z一-龥._()（） -]", "_", name).strip(" .")
        if not name or name in {".", ".."}:
            raise ValueError("上传文件名无效")
        return name

    def create_run(
        self,
        *,
        year: int,
        month: int,
        operator: str,
        uploads: Iterable[Any],
    ) -> dict[str, Any]:
        if not 1 <= month <= 12:
            raise ValueError("月份必须为1至12")
        operator = operator.strip()
        if not operator:
            raise ValueError("经办人不能为空")
        run_id = f"RH-{year:04d}-{month:02d}-{uuid4().hex[:8].upper()}"
        thread_id = uuid4().hex
        run_dir, input_dir = self._run_paths(run_id)
        input_dir.mkdir(parents=True, exist_ok=False)
        try:
            uploaded = self._write_uploads(input_dir, uploads, allow_overwrite=False)
            if not uploaded:
                raise ValueError("至少上传一个资料文件")
        except Exception:
            shutil.rmtree(run_dir, ignore_errors=True)
            raise
        (run_dir / "upload_manifest.json").write_text(
            json.dumps(
                {"run_id": run_id, "thread_id": thread_id, "uploaded": uploaded},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
            )
        row = self.app_repository.create_run(
            run_id=run_id,
            thread_id=thread_id,
            year=year,
            month=month,
            operator=operator,
            source_dir=input_dir,
            output_root=run_dir / "output",
        )
        (run_dir / "metadata.json").write_text(
            json.dumps(
                {
                    "run_id": run_id,
                    "thread_id": thread_id,
                    "year": year,
                    "month": month,
                    "operator": operator,
                    "input_dir": str(input_dir),
                    "output_root": str(run_dir / "output"),
                    "uploaded": uploaded,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        row["uploaded"] = uploaded
        return row

    def add_uploads(self, run_id: str, uploads: Iterable[Any]) -> dict[str, Any]:
        """在启动前补充或替换资料，并使之前的预检结果失效。"""
        row = self._require_run(run_id)
        if row["status"] != "DRAFT":
            raise ValueError("只有DRAFT任务可以补充或替换资料")
        _, input_dir = self._run_paths(run_id)
        uploaded = self._write_uploads(input_dir, uploads, allow_overwrite=True)
        if not uploaded:
            raise ValueError("至少上传一个资料文件")
        manifest_path = input_dir.parent / "upload_manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.is_file() else {}
        names = sorted({*manifest.get("uploaded", []), *uploaded})
        manifest.update({"run_id": run_id, "thread_id": row["thread_id"], "uploaded": names})
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        self.app_repository.update(
            run_id,
            preflight_status="NOT_RUN",
            preflight_result=None,
            error_message=None,
        )
        result = self._require_run(run_id)
        result["uploaded"] = uploaded
        return result

    def _preflight_state(self, row: dict[str, Any]) -> dict[str, Any]:
        from langgraph_profit_workflow import determine_period_policy, preflight

        base = {
            "year": row["year"],
            "month": row["month"],
            "source_dir": row["source_dir"],
            "output_root": row["output_root"],
        }
        policy_state = {**base, **determine_period_policy(base)}
        return {**policy_state, **preflight(policy_state)}

    def preflight(self, run_id: str) -> dict[str, Any]:
        row = self._require_run(run_id)
        if row["status"] not in {"DRAFT", "FAILED"}:
            raise ValueError(f"当前状态不能预检：{row['status']}")
        self.app_repository.update(run_id, status="VALIDATING", error_message=None)
        try:
            state = self._preflight_state(row)
            selected = state["selected_files"]
            files = [
                {"role": key, "path": path}
                for key, value in selected.items()
                for path in (value if isinstance(value, list) else [value])
            ]
            payload = {
                "ok": True,
                "files": files,
                "selected_files": selected,
                "current_input_fingerprint": state["current_input_fingerprint"],
                "prior_input_fingerprint": state["prior_input_fingerprint"],
            }
            run_dir, _ = self._run_paths(run_id)
            (run_dir / "preflight.json").write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            self.app_repository.update(
                run_id,
                status="DRAFT",
                preflight_status="PASSED",
                preflight_result=json.dumps(payload, ensure_ascii=False),
                error_message=None,
            )
        except Exception as error:
            payload = {"ok": False, "error": f"{type(error).__name__}: {error}"}
            self.app_repository.update(
                run_id,
                status="DRAFT",
                preflight_status="FAILED",
                preflight_result=json.dumps(payload, ensure_ascii=False),
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
        self.app_repository.update(run_id, status="QUEUED", error_message=None)
        self._executor.submit(self._invoke_start, row)
        return self.get_status(run_id)

    def resume(self, run_id: str, correction: dict[str, Any]) -> dict[str, Any]:
        row = self._require_run(run_id)
        if row["status"] != "WAITING_REVIEW":
            raise ValueError(f"当前状态不能恢复：{row['status']}")
        self.app_repository.update(run_id, status="QUEUED", error_message=None)
        self._executor.submit(self._invoke_resume, row, correction)
        return self.get_status(run_id)

    def _config(self, row: dict[str, Any]) -> dict[str, Any]:
        return {"configurable": {"thread_id": row["thread_id"]}}

    def _invoke_start(self, row: dict[str, Any]) -> None:
        self.app_repository.update(row["run_id"], status="RUNNING")
        try:
            from langgraph_profit_workflow import ProfitWorkflowInput

            workflow_input = ProfitWorkflowInput(
                year=row["year"],
                month=row["month"],
                source_dir=Path(row["source_dir"]),
                output_root=Path(row["output_root"]),
            )
            with self._graph_lock:
                result = self._ensure_graph().invoke(
                    workflow_input.model_dump(mode="json"), config=self._config(row)
                )
            self._sync_result(row, result)
        except Exception as error:
            self._record_failure(row, error)

    def _invoke_resume(self, row: dict[str, Any], correction: dict[str, Any]) -> None:
        self.app_repository.update(row["run_id"], status="RUNNING")
        try:
            from langgraph.types import Command

            with self._graph_lock:
                result = self._ensure_graph().invoke(
                    Command(resume=correction), config=self._config(row)
                )
            self._sync_result(row, result)
        except Exception as error:
            self._record_failure(row, error)

    def _sync_result(self, row: dict[str, Any], result: Any) -> None:
        state = self._state(row)
        if isinstance(result, dict) and result.get("__interrupt__"):
            status = "WAITING_REVIEW"
        else:
            status = state.get("workflow_status", result.get("workflow_status", "RUNNING"))
        updates: dict[str, Any] = {"status": status}
        if state.get("run_output_dir"):
            updates["run_output_dir"] = state["run_output_dir"]
        if state.get("formal_report"):
            updates["formal_report"] = state["formal_report"]
        if status in {"SUCCEEDED", "FAILED"}:
            updates["completed_at"] = datetime.now().isoformat(timespec="seconds")
        self.app_repository.update(row["run_id"], **updates)

    def _record_failure(self, row: dict[str, Any], error: Exception) -> None:
        self.app_repository.update(
            row["run_id"],
            status="FAILED",
            error_message=f"{type(error).__name__}: {error}",
            completed_at=datetime.now().isoformat(timespec="seconds"),
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
        # 工作线程执行 graph.invoke 时持有 graph 锁；运行中状态直接读业务索引，
        # 避免前端轮询被整段 Excel 计算阻塞。
        state = self._state(row) if row["status"] in {"WAITING_REVIEW", "SUCCEEDED", "FAILED"} else {}
        artifacts = state.get("artifacts", {})
        preflight = None
        if row.get("preflight_result"):
            try:
                preflight = json.loads(row["preflight_result"])
            except (TypeError, json.JSONDecodeError):
                preflight = {"ok": False, "error": row.get("preflight_result")}
        return {
            **row,
            "preflight": preflight,
            "workflow_status": state.get("workflow_status", row["status"]),
            "review_round": state.get("review_round", 0),
            "m7_differences": state.get("m7_differences", []),
            "subject_differences": state.get("subject_differences", []),
            "unbilled_evidence": state.get("unbilled_evidence", []),
            "difference_review_path": state.get("difference_review_path"),
            "artifacts": {key: Path(value).name for key, value in artifacts.items()},
            "formal_report": state.get("formal_report", row.get("formal_report")),
            "error": state.get("error"),
        }

    def list_runs(self) -> list[dict[str, Any]]:
        return self.app_repository.list()

    def artifact_path(self, run_id: str, key: str) -> Path:
        row = self._require_run(run_id)
        state = self._state(row)
        path_text = state.get("artifacts", {}).get(key)
        if not path_text:
            raise FileNotFoundError(f"产物不存在：{key}")
        path = Path(path_text).resolve()
        run_dir = self.runs_root.joinpath(run_id).resolve()
        if run_dir not in path.parents or not path.is_file():
            raise FileNotFoundError("产物路径无效")
        return path

    def _require_run(self, run_id: str) -> dict[str, Any]:
        row = self.app_repository.get(run_id)
        if row is None:
            raise KeyError(f"运行不存在：{run_id}")
        return row

    def close(self) -> None:
        self._executor.shutdown(wait=True)
        if self._connection is not None:
            self._connection.close()
