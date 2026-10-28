import hashlib
import io
import tempfile
import unittest
from pathlib import Path
from uuid import uuid4

import sys

SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from renheng_web.database import RunRepository
from renheng_web.app import create_app
from renheng_web.security import hash_password
from renheng_web.service import WorkflowService
from renheng_web.settings import Settings
from fastapi.testclient import TestClient


class FakeUpload:
    def __init__(self, filename: str, content: bytes):
        self.filename = filename
        self.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        self.file = io.BytesIO(content)


class FakeObjectResponse:
    def __init__(self, content: bytes):
        self._stream = io.BytesIO(content)

    def stream(self, chunk_size: int):
        while True:
            chunk = self._stream.read(chunk_size)
            if not chunk:
                break
            yield chunk

    def read(self, size: int = -1):
        return self._stream.read(size)

    def close(self):
        self._stream.close()

    def release_conn(self):
        pass


class FakeObjectStore:
    def __init__(self):
        self.objects: dict[str, bytes] = {}
        self.bucket = "renheng-profit"

    def put_upload(self, upload, object_key):
        upload.file.seek(0)
        content = upload.file.read()
        upload.file.seek(0)
        self.objects[object_key] = content
        return {
            "size_bytes": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
            "content_type": upload.content_type,
        }

    def put_path(self, path, object_key, content_type):
        content = path.read_bytes()
        self.objects[object_key] = content
        return {
            "size_bytes": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
            "content_type": content_type,
        }

    def download_to_path(self, object_key, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(self.objects[object_key])

    def open_download(self, object_key):
        return FakeObjectResponse(self.objects[object_key])

    def delete(self, object_key):
        self.objects.pop(object_key, None)


class WebServiceTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        root = Path(self.temp_dir.name)
        self.repository = RunRepository(f"sqlite:///{root / 'test.sqlite'}", create_schema=True)
        self.user = self.repository.create_user(
            username="finance",
            display_name="财务",
            role="FINANCE",
            password_hash="pbkdf2_sha256$1$a$b",
        )
        self.reviewer = self.repository.create_user(
            username="reviewer",
            display_name="审核人",
            role="REVIEWER",
            password_hash="pbkdf2_sha256$1$a$b",
        )
        self.settings = Settings(
            database_url="sqlite://",
            checkpoint_database_url="postgresql://unused",
            minio_endpoint="localhost:9000",
            minio_access_key="test",
            minio_secret_key="test-secret",
            session_secret_key="x" * 48,
        )
        self.storage = FakeObjectStore()
        self.service = WorkflowService(
            root / "data",
            settings=self.settings,
            repository=self.repository,
            storage=self.storage,
            recover_pending=False,
        )

    def tearDown(self):
        self.service.close()
        self.temp_dir.cleanup()

    def test_run_repository_persists_distinct_run_and_checkpoint_ids(self):
        row = self.repository.create_run(
            run_id="RH-2026-02-TEST",
            thread_id="thread-001",
            year=2026,
            month=2,
            created_by=self.user["user_id"],
        )
        self.repository.update(row["run_id"], status="WAITING_REVIEW")
        saved = self.repository.get(row["run_id"])
        self.assertEqual(saved["thread_id"], "thread-001")
        self.assertNotEqual(saved["run_id"], saved["thread_id"])
        self.assertEqual(saved["status"], "WAITING_REVIEW")
        self.assertEqual(saved["operator"], "财务")

    def test_interrupted_upload_can_be_retried_after_restart(self):
        row = self.repository.create_run(
            run_id="RH-2026-05-UPLOAD",
            thread_id="thread-upload",
            year=2026,
            month=5,
            created_by=self.user["user_id"],
        )
        self.assertEqual(row["status"], "UPLOADING")
        self.assertEqual(self.repository.fail_interrupted_uploads(), 1)
        failed = self.repository.get(row["run_id"])
        self.assertEqual(failed["status"], "UPLOAD_FAILED")

        recovered = self.service.add_uploads(
            row["run_id"], [FakeUpload("利润表.xlsx", b"reuploaded")], user=self.user
        )
        self.assertEqual(recovered["status"], "DRAFT")
        self.assertEqual(recovered["preflight_status"], "NOT_RUN")

    def test_failed_artifact_index_write_removes_minio_object(self):
        from unittest.mock import patch

        row = self.service.create_run(
            year=2026,
            month=2,
            user=self.user,
            uploads=[FakeUpload("利润表.xlsx", b"workbook-content")],
        )
        _, _, output_dir = self.service._run_paths(row["run_id"])
        artifact_path = output_dir / "M1" / "manifest.json"
        artifact_path.parent.mkdir(parents=True)
        artifact_path.write_text("{}", encoding="utf-8")
        with patch.object(self.repository, "add_file", side_effect=RuntimeError("database unavailable")):
            with self.assertRaisesRegex(RuntimeError, "database unavailable"):
                self.service._persist_local_file(
                    row["run_id"],
                    artifact_key="m1_input_manifest",
                    path=artifact_path,
                    created_by=self.user["user_id"],
                    role="ARTIFACT",
                )
        self.assertEqual(len(self.storage.objects), 1)

    def test_upload_bytes_are_in_object_store_and_file_index(self):
        row = self.service.create_run(
            year=2026,
            month=2,
            user=self.user,
            uploads=[FakeUpload("利润表.xlsx", b"workbook-content")],
        )
        self.assertEqual(row["status"], "DRAFT")
        self.assertNotEqual(row["run_id"], row["thread_id"])
        files = self.repository.list_files(row["run_id"])
        self.assertEqual(len(files), 1)
        self.assertEqual(files[0]["role"], "INPUT")
        self.assertIn(files[0]["object_key"], self.storage.objects)
        path = Path(self.service._run_paths(row["run_id"])[1]) / "利润表.xlsx"
        self.assertEqual(path.read_bytes(), b"workbook-content")
        self.assertEqual(files[0]["sha256"], hashlib.sha256(b"workbook-content").hexdigest())

    def test_replacing_input_keeps_version_history_and_invalidates_preflight(self):
        row = self.service.create_run(
            year=2026,
            month=2,
            user=self.user,
            uploads=[FakeUpload("资料.xlsx", b"old")],
        )
        self.repository.update(row["run_id"], preflight_status="PASSED", preflight_result={"ok": True})
        updated = self.service.add_uploads(
            row["run_id"], [FakeUpload("资料.xlsx", b"new")], user=self.user
        )
        files = self.repository.list_files(row["run_id"], current_only=False)
        current = self.repository.list_files(row["run_id"], current_only=True)
        self.assertEqual(updated["preflight_status"], "NOT_RUN")
        self.assertEqual(len(files), 2)
        self.assertEqual(len(current), 1)
        self.assertEqual(current[0]["version"], 2)
        self.assertEqual(self.storage.objects[current[0]["object_key"]], b"new")

    def test_review_record_and_queue_transition_are_atomic_and_append_only(self):
        row = self.repository.create_run(
            run_id="RH-2026-03-TEST",
            thread_id="thread-003",
            year=2026,
            month=3,
            created_by=self.user["user_id"],
        )
        self.repository.update(row["run_id"], status="WAITING_REVIEW")
        self.repository.create_approval_and_queue(
            run_id=row["run_id"],
            review_round=1,
            reviewer_id=self.reviewer["user_id"],
            reason="台账核实",
            corrected_revenue={"current_month": "100.00"},
            before_differences=[{"差异": "10.00"}],
        )
        self.assertEqual(self.repository.get(row["run_id"])["status"], "QUEUED")
        approvals = self.repository.list_approvals(row["run_id"])
        self.assertEqual(len(approvals), 1)
        self.assertEqual(approvals[0]["reviewer"], "审核人")
        self.assertEqual(approvals[0]["corrected_revenue"]["current_month"], "100.00")
        with self.assertRaises(ValueError):
            self.repository.create_approval_and_queue(
                run_id=row["run_id"],
                review_round=1,
                reviewer_id=self.reviewer["user_id"],
                reason="重复提交",
                corrected_revenue={},
                before_differences=[],
            )

    def test_status_returns_file_index_not_local_paths(self):
        row = self.service.create_run(
            year=2026,
            month=2,
            user=self.user,
            uploads=[FakeUpload("利润表.xlsx", b"workbook-content")],
        )
        result = self.service.get_status(row["run_id"])
        self.assertNotIn("source_dir", result)
        self.assertEqual(result["files"][0]["role"], "INPUT")

    def test_workflow_artifact_checkpoint_value_is_object_ref_and_hydrates_on_resume(self):
        row = self.service.create_run(
            year=2026,
            month=2,
            user=self.user,
            uploads=[FakeUpload("利润表.xlsx", b"workbook-content")],
        )
        run_dir, _, output_dir = self.service._run_paths(row["run_id"])
        artifact_path = output_dir / "M1" / "manifest.json"
        artifact_path.parent.mkdir(parents=True)
        artifact_path.write_text('{"thread_id":"stable"}', encoding="utf-8")
        updates = self.service._persist_node_updates(
            row["run_id"], {"artifacts": {"m1_input_manifest": str(artifact_path)}}
        )
        reference = updates["artifacts"]["m1_input_manifest"]
        self.assertTrue(reference.startswith("minio://renheng-profit/runs/"))
        record = self.repository.get_artifact(row["run_id"], "m1_input_manifest")
        self.assertIsNotNone(record)
        self.assertIn(record["object_key"], self.storage.objects)
        artifact_path.unlink()
        restored = self.service._resolve_artifact_reference(row["run_id"], reference)
        self.assertEqual(restored.read_text(encoding="utf-8"), '{"thread_id":"stable"}')

    def test_minio_reference_cannot_cross_run_boundary(self):
        row = self.service.create_run(
            year=2026,
            month=2,
            user=self.user,
            uploads=[FakeUpload("利润表.xlsx", b"workbook-content")],
        )
        with self.assertRaises(FileNotFoundError):
            self.service._resolve_artifact_reference(
                row["run_id"], "minio://renheng-profit/runs/another-run/artifacts/file"
            )

    def test_login_is_required_and_finance_user_cannot_submit_review(self):
        finance = self.repository.create_user(
            username="login-finance",
            display_name="登录财务",
            role="FINANCE",
            password_hash=hash_password("finance password with length"),
        )
        waiting = self.repository.create_run(
            run_id="RH-2026-04-WAIT",
            thread_id="thread-waiting",
            year=2026,
            month=4,
            created_by=finance["user_id"],
        )
        self.repository.update(waiting["run_id"], status="WAITING_REVIEW")
        app = create_app(Path(self.temp_dir.name) / "data", service=self.service)
        with TestClient(app) as client:
            self.assertEqual(client.get("/api/runs").status_code, 401)
            login = client.post(
                "/api/auth/login",
                json={"username": "login-finance", "password": "finance password with length"},
            )
            self.assertEqual(login.status_code, 200)
            self.assertEqual(client.get("/api/auth/me").json()["role"], "FINANCE")
            resume = client.post(
                f"/api/runs/{waiting['run_id']}/resume",
                json={"reason": "审核", "corrected_revenue": {"current_month": "1.00"}},
            )
            self.assertEqual(resume.status_code, 403)


if __name__ == "__main__":
    unittest.main()
