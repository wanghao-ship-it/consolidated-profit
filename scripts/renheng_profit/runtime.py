"""输入指纹、跨模块预检和运行状态持久化。"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any, Iterable

import xlrd

from .common import SourceFile, json_default
from .income_statements import parse_income_statement


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def input_fingerprint(paths: Iterable[Path]) -> str:
    digest = hashlib.sha256()
    for path in sorted({Path(path).resolve() for path in paths}, key=lambda value: str(value).casefold()):
        digest.update(str(path).encode("utf-8"))
        digest.update(sha256_file(path).encode("ascii"))
    return digest.hexdigest()


def validate_template(template_file: Path) -> None:
    if not template_file.is_file():
        raise FileNotFoundError(f"模板文件不存在：{template_file}")
    if template_file.suffix.lower() != ".xls":
        raise ValueError("模板文件后缀必须严格为.xls")
    workbook = xlrd.open_workbook(template_file)
    if workbook.nsheets < 2:
        raise ValueError("模板至少需要累计表和本月表两个Sheet")


def validate_profit_sources(sources: list[SourceFile]) -> None:
    for source in sources:
        parse_income_statement(source)


def write_run_status(output_dir: Path, payload: dict[str, Any]) -> Path:
    path = output_dir / "run_status.json"
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=json_default), encoding="utf-8"
    )
    return path
