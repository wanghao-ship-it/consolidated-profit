"""共享常量、数据模型和无副作用基础工具。"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
import re

RULE_VERSION = "2026-09-langgraph-zero-difference-v4"
MONEY_QUANTUM = Decimal("0.01")
ZERO = Decimal("0.00")
COMPANIES = ("仁恒科技", "阳原仁恒", "上海分公司", "仁恒化工", "新疆公司", "美国分公司")
COMPANY_FILE_KEYWORDS = {
    "仁恒科技": ("仁恒科技",),
    "阳原仁恒": ("阳原仁恒",),
    "上海分公司": ("上海分公司",),
    "仁恒化工": ("仁恒化工", "化工"),
    "新疆公司": ("新疆",),
    "美国分公司": ("美国", "Clayminton"),
}
US_COMPANY = "美国分公司"
QUARTER_END_MONTHS = frozenset({3, 6, 9, 12})
SUMMARY_MONTHLY_COLUMN = "六主体汇总_本月"
SUMMARY_CUMULATIVE_COLUMN = "六主体汇总_累计"
PRE_SUMMARY_FILENAME = "03_六主体利润表汇总_抵销前.csv"
ITEM_ALIASES = {
    "营业收入": ("营业收入",), "营业成本": ("营业成本",),
    "税金及附加": ("税金及附加",), "销售费用": ("销售费用",),
    "管理费用": ("管理费用",), "研发费用": ("研发费用", "研究费用"),
    "财务费用": ("财务费用",), "其他收益": ("其他收益",),
    "资产减值损失": ("资产减值损失",),
    "公允价值变动收益": ("公允价值变动收益",),
    "投资收益": ("投资收益",), "营业利润": ("营业利润",),
    "营业外收入": ("营业外收入",), "营业外支出": ("营业外支出",),
    "以前年度损益调整": ("以前年度损益调整",),
    "利润总额": ("利润总额",), "所得税费用": ("所得税费用",),
    "净利润": ("净利润",),
}
KEY_ITEMS = ("营业收入", "营业成本", "营业利润", "利润总额", "净利润")
TEMPLATE_ITEM_ROWS = {
    "营业收入": 4,
    "营业成本": 5,
    "销售费用": 6,
    "管理费用": 7,
    "研发费用": 8,
    "财务费用": 9,
    "资产减值损失": 10,
    "税金及附加": 11,
    "公允价值变动收益": 12,
    "投资收益": 13,
    "营业利润": 15,
    "营业外收入": 16,
    "营业外支出": 17,
    "以前年度损益调整": 18,
    "利润总额": 19,
    "所得税费用": 20,
    "净利润": 22,
}
TEMPLATE_BLANK_ZERO_ITEMS = {"资产减值损失", "公允价值变动收益", "以前年度损益调整"}
M6_METRIC_ROWS = {
    "毛利率": 28,
    "毛利": 29,
    "营业净利率": 35,
    "税前利率": 36,
    "主营业务成本率": 37,
}
M6_METRIC_ORDER = tuple(M6_METRIC_ROWS)
M6_NO_CHANGE_METRICS = {"营业净利率", "税前利率", "主营业务成本率"}
GROUP_COMPANIES = {
    "科技": "浙江长安仁恒科技股份有限公司",
    "阳原": "阳原县仁恒精细粘土有限责任公司",
    "上海": "浙江长安仁恒科技股份有限公司上海分公司",
    "仁恒化工": "浙江长安仁恒化工有限公司",
    "新疆公司": "长安仁恒（新疆）科技技术有限公司",
    "美国分公司": "Clayminton INC",
}
GROUP_COMPANY_NAMES = frozenset(GROUP_COMPANIES.values())
SOURCE_CONFIG = {
    "科技": {
        "file_keywords": ("科技",),
        "sheet_names": ("开票基础信息", "发票基础信息"),
        "seller_name": GROUP_COMPANIES["科技"],
    },
    "阳原": {
        "file_keywords": ("阳原",),
        "sheet_names": ("发票基础信息",),
        "seller_name": GROUP_COMPANIES["阳原"],
    },
    "上海": {
        "file_keywords": ("上海",),
        "sheet_names": ("开票基础信息", "发票基础信息"),
        "seller_name": GROUP_COMPANIES["上海"],
    },
    "化工": {
        "file_keywords": ("化工",),
        "sheet_names": ("开票基础信息", "发票基础信息"),
        "seller_name": GROUP_COMPANIES["仁恒化工"],
    },
    "新疆": {
        "file_keywords": ("新疆",),
        "sheet_names": ("开票基础信息", "发票基础信息"),
        "seller_name": GROUP_COMPANIES["新疆公司"],
    },
    "美国": {
        "file_keywords": ("美国", "Clayminton"),
        "sheet_names": ("开票基础信息", "发票基础信息"),
        "seller_name": GROUP_COMPANIES["美国分公司"],
    },
}
REQUIRED_INVOICE_SOURCES = frozenset({"科技", "阳原"})
INVOICE_REVENUE_COMPANY_MAP = {
    "科技": "仁恒科技",
    "阳原": "阳原仁恒",
}
HEADER_ALIASES = {
    "销方名称": ("销方名称",),
    "购买方名称": ("购买方名称",),
    "开票日期": ("开票日期",),
    # 科技文件使用“金额”，阳原文件使用“不含税金额”，两者均为本脚本的内部销售金额口径。
    "金额": ("金额", "不含税金额"),
}
INVOICE_LINK_FIELDS = ("数电发票号码", "发票代码", "发票号码")
ELIMINATION_ITEMS = {"营业收入"}
OPERATING_COST_COMPONENT_WEIGHTS = {
    "税金及附加": Decimal("1"),
    "销售费用": Decimal("1"),
    "管理费用": Decimal("1"),
    "研发费用": Decimal("1"),
    "财务费用": Decimal("1"),
    # 投资收益、公允价值变动收益计入营业利润（源表已含），倒挤时须反向扣除：
    # 营业成本 = 营业收入 -（税金+费用 - 投资收益 - 公允价值 + 营业利润）
    "投资收益": Decimal("-1"),
    "营业利润": Decimal("1"),
    "公允价值变动收益": Decimal("-1"),
}
REQUIRED_PRE_COLUMNS = {"项目", SUMMARY_MONTHLY_COLUMN, SUMMARY_CUMULATIVE_COLUMN}
REQUIRED_INTERNAL_COLUMNS = {"范围", "销方主体", "购买方名称", "内部销售金额"}
LEDGER_HEADER_ALIASES = {
    "月份": ("月份",),
    "销方名称": ("销方名称",),
    "金额（不含税）": ("金额(不含税)", "金额（不含税）"),
}
LEDGER_BUYER_ALIASES = ("购买方名称", "购方名称")
LEDGER_INVOICE_NUMBER_ALIASES = ("数电票号码", "数电发票号码")


@dataclass
class SourceFile:
    company: str
    path: str
    suffix: str
    size_bytes: int


@dataclass
class ParseMetadata:
    company: str
    file: str
    sheet: str
    header_row: int
    project_column: int
    monthly_column: int
    cumulative_column: int
    source_row_count: int
    recognized_item_count: int


@dataclass
class InternalSaleRecord:
    period: str
    source_company: str
    seller_name: str
    buyer_name: str
    invoice_date: str
    amount: Decimal
    source_file: str
    source_sheet: str
    source_row: int
    invoice_status: str


@dataclass
class LedgerDataset:
    file: Path
    sheet: str
    header_row: int
    has_invoice_number: bool
    rows: list[dict[str, Any]]


def compact_text(value: Any) -> str:
    return "" if value is None else re.sub(r"[\s\u3000]+", "", str(value)).strip()


def to_number(value: Any) -> Decimal | None:
    """将源表金额统一为两位小数的十进制金额，避免 float 尾数误差。"""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        text = str(value)
    else:
        text = str(value).strip().replace(",", "")
    if text in {"", "-", "--"}:
        return None
    try:
        return Decimal(text).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def json_default(value: Any) -> str:
    """JSON 中以固定两位小数保存金额，保证结果可读且无浮点尾数。"""
    if isinstance(value, Decimal):
        return format(value, ".2f")
    raise TypeError(f"无法序列化类型：{type(value).__name__}")


def target_month_from_period_token(period_token: str | None) -> int | None:
    if not period_token:
        return None
    match = re.search(r"20\d{2}[.年](\d{1,2})", compact_text(period_token))
    return None if match is None else int(match.group(1))


def parse_profit_period(file_name: str, report_year: str) -> int | None:
    """从利润表文件名解析截止月；季度名和年月名统一返回1至12月。"""
    normalized = compact_text(file_name)
    if report_year not in normalized:
        return None
    lowered = normalized.casefold()
    quarter_patterns = {
        3: ("q1", "第一季度", "第1季度", "1季度"),
        6: ("q2", "第二季度", "第2季度", "2季度"),
        9: ("q3", "第三季度", "第3季度", "3季度"),
        12: ("q4", "第四季度", "第4季度", "4季度"),
    }
    for end_month, tokens in quarter_patterns.items():
        if any(token in lowered for token in tokens):
            return end_month
    match = re.search(
        rf"{re.escape(report_year)}(?:年|[.\-_/])[^0-9]{{0,40}}0?(1[0-2]|[1-9])(?:月)?(?!\d)",
        normalized,
    )
    return int(match.group(1)) if match else None


def ignore_us_profit_for_period(report_year: str | int, target_month: int | None) -> bool:
    """2025年美国分公司从6月开始有数据；1至5月不识别、不警告、不审核。"""
    return str(report_year) == "2025" and target_month is not None and target_month < 6


def active_group_company_names(include_us_company: bool) -> frozenset[str]:
    return frozenset(
        company_name
        for company, company_name in GROUP_COMPANIES.items()
        if include_us_company or company != US_COMPANY
    )


def is_group_company(
    value: Any, group_company_names: frozenset[str] = GROUP_COMPANY_NAMES
) -> bool:
    """集团主体名称比较忽略英文大小写和前后空格。"""
    normalized_names = frozenset(
        compact_text(company_name).casefold() for company_name in group_company_names
    )
    return compact_text(value).casefold() in normalized_names


def invoice_to_money(value: Any) -> Decimal:
    if value is None or value == "":
        return ZERO
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "--"}:
        return ZERO
    try:
        return Decimal(text).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"金额列无法识别为金额：{value!r}") from error


def parse_invoice_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_file_period(file_name: str) -> tuple[int, int] | None:
    """兼容2026.2月、2026年2月、2025年仁恒科技2月等文件名。"""
    match = re.search(
        r"(?P<year>20\d{2})[.年][^0-9]{0,30}(?P<month>\d{1,2})月?",
        compact_text(file_name),
    )
    if not match:
        return None
    return int(match.group("year")), int(match.group("month"))


def normalized_invoice_key(value: Any) -> str:
    return compact_text(value).casefold()


def elimination_to_money(value: Any) -> Decimal:
    if value is None or value == "":
        return ZERO
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "--"}:
        return ZERO
    try:
        return Decimal(text).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"无法识别金额：{value!r}") from error


def elimination_json_default(value: Any) -> str:
    if isinstance(value, Decimal):
        return format(value, ".2f")
    raise TypeError(f"无法序列化类型：{type(value).__name__}")


def metric_to_text(value: Decimal | None, is_amount: bool = False) -> str:
    if value is None:
        return ""
    return format(value, ".2f" if is_amount else ".8f")
