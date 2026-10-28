import io
import json
import tempfile
import unittest
from pathlib import Path

import sys

SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from renheng_web.database import RunRepository
from renheng_web.service import WorkflowService


class FakeUpload:
    def __init__(self, filename: str, content: bytes):
        self.filename = filename
        self.file = io.BytesIO(content)


class WebServiceTests(unittest.TestCase):
    def test_run_repository_persists_thread_id_and_status(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            repository = RunRepository(Path(temp_dir) / "app.sqlite")
            repository.create_run(
                run_id="RH-2026-02-TEST",
                thread_id="thread-001",
                year=2026,
                month=2,
                operator="财务",
                source_dir=Path(temp_dir) / "input",
                output_root=Path(temp_dir),
            )
            repository.update("RH-2026-02-TEST", status="WAITING_REVIEW")
            row = repository.get("RH-2026-02-TEST")

        self.assertEqual(row["thread_id"], "thread-001")
        self.assertEqual(row["status"], "WAITING_REVIEW")

    def test_create_run_isolates_uploaded_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            service = WorkflowService(Path(temp_dir))
            try:
                self.assertTrue((Path(temp_dir) / "langgraph_checkpoints.sqlite").is_file())
                row = service.create_run(
                    year=2026,
                    month=2,
                    operator="财务",
                    uploads=[FakeUpload("利润表.xlsx", b"test-content")],
                )
                input_file = Path(row["source_dir"]) / "利润表.xlsx"
                manifest = Path(temp_dir) / "runs" / row["run_id"] / "upload_manifest.json"
                self.assertTrue(input_file.is_file())
                self.assertEqual(input_file.read_bytes(), b"test-content")
                self.assertEqual(json.loads(manifest.read_text(encoding="utf-8"))["uploaded"], ["利润表.xlsx"])
                self.assertNotEqual(row["run_id"], row["thread_id"])
                self.assertEqual(Path(row["output_root"]), Path(temp_dir) / "runs" / row["run_id"] / "output")
            finally:
                service.close()

    def test_preflight_keeps_incomplete_task_editable(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            service = WorkflowService(Path(temp_dir))
            try:
                row = service.create_run(
                    year=2026,
                    month=2,
                    operator="财务",
                    uploads=[FakeUpload("部分资料.xlsx", b"test-content")],
                )
                result = service.preflight(row["run_id"])
                self.assertEqual(result["status"], "DRAFT")
                self.assertEqual(result["preflight_status"], "FAILED")
                self.assertIn("preflight", result)
                self.assertFalse(result["preflight"]["ok"])
                replaced = service.add_uploads(
                    row["run_id"], [FakeUpload("部分资料.xlsx", b"new-content")]
                )
                self.assertEqual(replaced["preflight_status"], "NOT_RUN")
                self.assertEqual(
                    (Path(row["source_dir"]) / "部分资料.xlsx").read_bytes(), b"new-content"
                )
            finally:
                service.close()


if __name__ == "__main__":
    unittest.main()
