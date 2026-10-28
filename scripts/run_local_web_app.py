#!/usr/bin/env python3
"""启动单机版仁恒合并利润表网页工作台。"""
from __future__ import annotations

import argparse
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="启动仁恒合并利润表本地网页工作台")
    parser.add_argument("--data-root", type=Path, default=Path("data"), help="运行数据根目录")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址")
    parser.add_argument("--port", type=int, default=8000, help="监听端口")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        import uvicorn
        from renheng_web.app import create_app
        from renheng_web.settings import ConfigurationError
    except ImportError as error:
        raise SystemExit(
            "缺少网页依赖，请先执行：python -m pip install -r requirements.txt"
        ) from error
    try:
        app = create_app(args.data_root)
    except (ConfigurationError, RuntimeError) as error:
        raise SystemExit(f"服务初始化失败：{error}") from error
    uvicorn.run(app, host=args.host, port=args.port)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
