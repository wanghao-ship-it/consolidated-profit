"""运行服务所需的外部配置。"""
from __future__ import annotations

from dataclasses import dataclass
import os


class ConfigurationError(RuntimeError):
    pass


@dataclass(frozen=True)
class Settings:
    database_url: str
    checkpoint_database_url: str
    minio_endpoint: str
    minio_access_key: str
    minio_secret_key: str
    minio_bucket: str = "renheng-profit"
    minio_secure: bool = False
    session_secret_key: str = ""
    session_cookie_secure: bool = False
    session_max_age_seconds: int = 8 * 60 * 60

    @classmethod
    def from_env(cls) -> "Settings":
        try:
            from dotenv import load_dotenv

            load_dotenv()
        except ImportError:
            pass
        required = {
            "DATABASE_URL": os.getenv("DATABASE_URL", "").strip(),
            "CHECKPOINT_DATABASE_URL": os.getenv("CHECKPOINT_DATABASE_URL", "").strip(),
            "MINIO_ENDPOINT": os.getenv("MINIO_ENDPOINT", "").strip(),
            "MINIO_ACCESS_KEY": os.getenv("MINIO_ACCESS_KEY", "").strip(),
            "MINIO_SECRET_KEY": os.getenv("MINIO_SECRET_KEY", "").strip(),
            "SESSION_SECRET_KEY": os.getenv("SESSION_SECRET_KEY", "").strip(),
        }
        missing = [name for name, value in required.items() if not value]
        if missing:
            raise ConfigurationError(
                "缺少运行配置：" + ", ".join(missing) + "。请复制.env.example为.env并填写。"
            )
        return cls(
            database_url=required["DATABASE_URL"],
            checkpoint_database_url=required["CHECKPOINT_DATABASE_URL"],
            minio_endpoint=required["MINIO_ENDPOINT"],
            minio_access_key=required["MINIO_ACCESS_KEY"],
            minio_secret_key=required["MINIO_SECRET_KEY"],
            minio_bucket=os.getenv("MINIO_BUCKET", "renheng-profit").strip(),
            minio_secure=_env_bool("MINIO_SECURE", False),
            session_secret_key=required["SESSION_SECRET_KEY"],
            session_cookie_secure=_env_bool("SESSION_COOKIE_SECURE", False),
            session_max_age_seconds=int(os.getenv("SESSION_MAX_AGE_SECONDS", str(8 * 60 * 60))),
        )


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}
