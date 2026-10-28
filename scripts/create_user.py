#!/usr/bin/env python3
"""Create a local authenticated finance-workbench user."""
from __future__ import annotations

import argparse
import getpass
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))


def main() -> int:
    try:
        from dotenv import load_dotenv

        load_dotenv()
    except ImportError:
        pass
    parser = argparse.ArgumentParser(description="Create a Renheng workflow user")
    parser.add_argument("username")
    parser.add_argument("display_name")
    parser.add_argument("--role", choices=("FINANCE", "REVIEWER", "ADMIN"), default="FINANCE")
    args = parser.parse_args()
    database_url = os.getenv("DATABASE_URL", "").strip()
    if not database_url:
        parser.error("请先设置DATABASE_URL")
    password = getpass.getpass("Password (at least 12 characters): ")
    confirm = getpass.getpass("Confirm password: ")
    if password != confirm:
        parser.error("两次输入的密码不一致")

    from renheng_web.database import RunRepository
    from renheng_web.security import hash_password

    repository = RunRepository(database_url)
    try:
        user = repository.create_user(
            username=args.username,
            display_name=args.display_name,
            role=args.role,
            password_hash=hash_password(password),
        )
    finally:
        repository.close()
    print(f"已创建用户 {user['username']}（{user['display_name']}，{user['role']}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
