"""本地财务工作台的密码校验和签名会话。"""
from __future__ import annotations

import base64
import hashlib
import hmac
import secrets
from typing import Any

from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer


PBKDF2_ITERATIONS = 600_000
SESSION_COOKIE = "renheng_session"


def hash_password(password: str) -> str:
    if len(password) < 12:
        raise ValueError("密码至少需要12个字符")
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return "pbkdf2_sha256${}${}${}".format(
        PBKDF2_ITERATIONS,
        base64.urlsafe_b64encode(salt).decode("ascii"),
        base64.urlsafe_b64encode(digest).decode("ascii"),
    )


def verify_password(password: str, encoded: str) -> bool:
    try:
        algorithm, rounds, salt_text, digest_text = encoded.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        salt = base64.urlsafe_b64decode(salt_text.encode("ascii"))
        expected = base64.urlsafe_b64decode(digest_text.encode("ascii"))
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, int(rounds))
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError, UnicodeError):
        return False


class SessionTokens:
    def __init__(self, secret_key: str, max_age_seconds: int = 8 * 60 * 60) -> None:
        if len(secret_key) < 32:
            raise ValueError("SESSION_SECRET_KEY至少需要32个字符")
        self.serializer = URLSafeTimedSerializer(secret_key, salt="renheng-finance-session-v1")
        self.max_age_seconds = max_age_seconds

    def issue(self, user: dict[str, Any]) -> str:
        return self.serializer.dumps({"user_id": user["user_id"]})

    def read_user_id(self, token: str | None) -> str | None:
        if not token:
            return None
        try:
            payload = self.serializer.loads(token, max_age=self.max_age_seconds)
        except (BadSignature, SignatureExpired):
            return None
        user_id = payload.get("user_id") if isinstance(payload, dict) else None
        return user_id if isinstance(user_id, str) else None
