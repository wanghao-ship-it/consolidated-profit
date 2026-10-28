"""MinIO对象存储适配器；数据库仅保存对象键和文件元数据。"""
from __future__ import annotations

import hashlib
from pathlib import Path
import shutil
from typing import Any, BinaryIO


class MinioArtifactStore:
    def __init__(
        self,
        endpoint: str,
        access_key: str,
        secret_key: str,
        bucket: str,
        *,
        secure: bool = False,
        client: Any | None = None,
        create_bucket: bool = True,
    ) -> None:
        if client is None:
            try:
                from minio import Minio
            except ImportError as error:
                raise RuntimeError("缺少MinIO依赖，请执行 python -m pip install -r requirements.txt") from error
            client = Minio(endpoint, access_key=access_key, secret_key=secret_key, secure=secure)
        self.client = client
        self.bucket = bucket
        if create_bucket and not self.client.bucket_exists(bucket):
            self.client.make_bucket(bucket)

    def put_upload(self, upload: Any, object_key: str) -> dict[str, Any]:
        stream: BinaryIO = upload.file
        try:
            stream.seek(0, 2)
            size = stream.tell()
            stream.seek(0)
        except (AttributeError, OSError) as error:
            raise ValueError("上传文件必须支持可定位的文件流") from error
        if size <= 0:
            raise ValueError(f"上传文件为空：{getattr(upload, 'filename', '')}")
        digest = _sha256_stream(stream)
        stream.seek(0)
        self.client.put_object(
            self.bucket,
            object_key,
            stream,
            length=size,
            content_type=getattr(upload, "content_type", None) or "application/octet-stream",
        )
        return {
            "size_bytes": size,
            "sha256": digest,
            "content_type": getattr(upload, "content_type", None) or "application/octet-stream",
        }

    def put_path(self, path: Path, object_key: str, content_type: str) -> dict[str, Any]:
        if not path.is_file():
            raise FileNotFoundError(f"待上传产物不存在：{path}")
        size = path.stat().st_size
        digest = _sha256_path(path)
        self.client.fput_object(self.bucket, object_key, str(path), content_type=content_type)
        return {"size_bytes": size, "sha256": digest, "content_type": content_type}

    def download_to_path(self, object_key: str, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        response = self.client.get_object(self.bucket, object_key)
        try:
            with path.open("wb") as target:
                shutil.copyfileobj(response, target, length=1024 * 1024)
        finally:
            response.close()
            release = getattr(response, "release_conn", None)
            if callable(release):
                release()

    def open_download(self, object_key: str) -> Any:
        return self.client.get_object(self.bucket, object_key)

    def delete(self, object_key: str) -> None:
        self.client.remove_object(self.bucket, object_key)


def _sha256_stream(stream: BinaryIO) -> str:
    digest = hashlib.sha256()
    while True:
        chunk = stream.read(1024 * 1024)
        if not chunk:
            break
        digest.update(chunk)
    return digest.hexdigest()


def _sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while True:
            chunk = stream.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()
