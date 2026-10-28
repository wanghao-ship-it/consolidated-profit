import io
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from renheng_web.storage import MinioArtifactStore


class FakeMinioClient:
    def __init__(self):
        self.objects = {}

    def bucket_exists(self, bucket):
        return True

    def make_bucket(self, bucket):
        raise AssertionError("bucket already exists")

    def put_object(self, bucket, object_name, data, length, content_type):
        self.objects[object_name] = data.read(length)

    def fput_object(self, bucket, object_name, file_path, content_type):
        self.objects[object_name] = Path(file_path).read_bytes()

    def get_object(self, bucket, object_name):
        return io.BytesIO(self.objects[object_name])

    def remove_object(self, bucket, object_name):
        self.objects.pop(object_name, None)


class Upload:
    filename = "資料.xlsx"
    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def __init__(self, content):
        self.file = io.BytesIO(content)


class StorageTests(unittest.TestCase):
    def test_upload_hash_and_download_round_trip(self):
        client = FakeMinioClient()
        store = MinioArtifactStore("localhost:9000", "test", "secret", "private", client=client)
        metadata = store.put_upload(Upload(b"excel-content"), "runs/run/input/file.xlsx")
        self.assertEqual(metadata["size_bytes"], len(b"excel-content"))
        self.assertEqual(client.objects["runs/run/input/file.xlsx"], b"excel-content")
        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "nested" / "download.xlsx"
            store.download_to_path("runs/run/input/file.xlsx", target)
            self.assertEqual(target.read_bytes(), b"excel-content")


if __name__ == "__main__":
    unittest.main()
