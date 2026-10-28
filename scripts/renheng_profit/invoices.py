"""开票收入、设备票排除、内部销售和阳原自销。"""
from __future__ import annotations

import csv
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .common import (
    GROUP_COMPANY_NAMES,
    HEADER_ALIASES,
    INVOICE_LINK_FIELDS,
    INVOICE_REVENUE_COMPANY_MAP,
    REQUIRED_INVOICE_SOURCES,
    SOURCE_CONFIG,
    ZERO,
    InternalSaleRecord,
    compact_text,
    invoice_to_money,
    is_group_company,
    normalized_invoice_key,
    parse_file_period,
    parse_invoice_date,
)


def find_invoice_file(
    source_dir: Path, source_company: str, year: int, month: int, required: bool = True
) -> Path | None:
    keywords = SOURCE_CONFIG[source_company]["file_keywords"]
    matches = []
    for path in source_dir.rglob("*"):
        if not path.is_file() or path.suffix.lower() != ".xlsx" or path.name.startswith("~$"):
            continue
        file_name = compact_text(path.name)
        if not any(keyword in file_name for keyword in keywords):
            continue
        if "开票" not in file_name and "发票" not in file_name:
            continue
        if parse_file_period(file_name) == (year, month):
            matches.append(path)

    if not matches and not required:
        return None
    if len(matches) != 1:
        detail = "未找到" if not matches else f"找到{len(matches)}份"
        raise FileNotFoundError(
            f"{source_company}{year}年{month}月开票文件{detail}；请保证输入目录中只有一份有效文件"
        )
    return matches[0]


def find_invoice_sheet(workbook, source_company: str):
    wanted_names = SOURCE_CONFIG[source_company]["sheet_names"]
    for wanted_name in wanted_names:
        if wanted_name in workbook.sheetnames:
            return workbook[wanted_name]
    raise ValueError(
        f"{source_company}开票文件缺少票级基础信息工作表；需要：{'、'.join(wanted_names)}"
    )


def header_columns(sheet) -> dict[str, int]:
    headers = [compact_text(cell.value) for cell in sheet[1]]
    columns = {}
    for field_name, aliases in HEADER_ALIASES.items():
        column = next((headers.index(alias) + 1 for alias in aliases if alias in headers), None)
        if column is None:
            raise ValueError(f"工作表“{sheet.title}”缺少必需列：{field_name}（可接受：{'、'.join(aliases)}）")
        columns[field_name] = column
    columns["发票状态"] = headers.index("发票状态") + 1 if "发票状态" in headers else 0
    return columns


def find_equipment_invoice_keys(workbook, file_path: Path) -> set[tuple[str, str]]:
    """从信息汇总表识别货物名称含“设备”的发票。"""
    if "信息汇总表" not in workbook.sheetnames:
        raise ValueError(f"{file_path.name}缺少“信息汇总表”，无法排除设备开票数据")
    summary_sheet = workbook["信息汇总表"]
    headers = {compact_text(cell.value): index + 1 for index, cell in enumerate(summary_sheet[1])}
    goods_column = headers.get("货物或应税劳务名称")
    if goods_column is None:
        raise ValueError(f"{file_path.name}的“信息汇总表”缺少“货物或应税劳务名称”列")
    link_fields = [field for field in INVOICE_LINK_FIELDS if field in headers]
    if not link_fields:
        raise ValueError(f"{file_path.name}的“信息汇总表”缺少发票关联字段：{'、'.join(INVOICE_LINK_FIELDS)}")

    equipment_keys: set[tuple[str, str]] = set()
    for row_index in range(2, summary_sheet.max_row + 1):
        goods_name = compact_text(summary_sheet.cell(row_index, goods_column).value)
        if "设备" not in goods_name:
            continue
        for field in link_fields:
            key = normalized_invoice_key(summary_sheet.cell(row_index, headers[field]).value)
            if key:
                equipment_keys.add((field, key))
    return equipment_keys


def is_equipment_invoice(
    sheet,
    row_index: int,
    base_headers: dict[str, int],
    equipment_keys: set[tuple[str, str]],
) -> bool:
    for field in INVOICE_LINK_FIELDS:
        if field not in base_headers:
            continue
        key = normalized_invoice_key(sheet.cell(row_index, base_headers[field]).value)
        if key and (field, key) in equipment_keys:
            return True
    return False


def calculate_invoice_revenue_overrides(
    source_dir: Path, year: int, month: int
) -> dict[str, dict[str, Decimal]]:
    """本年科技、阳原营业收入改取开票总额扣除设备发票后的金额。"""
    overrides: dict[str, dict[str, Decimal]] = {}
    for source_company, profit_company in INVOICE_REVENUE_COMPANY_MAP.items():
        monthly_invoice_total = ZERO
        monthly_equipment_total = ZERO
        cumulative_invoice_total = ZERO
        cumulative_equipment_total = ZERO
        config = SOURCE_CONFIG[source_company]

        for period_month in range(1, month + 1):
            invoice_file = find_invoice_file(source_dir, source_company, year, period_month)
            if invoice_file is None:
                raise RuntimeError(f"{source_company}{year}年{period_month}月开票文件未找到")
            workbook = load_workbook(invoice_file, read_only=False, data_only=True)
            sheet = find_invoice_sheet(workbook, source_company)
            columns = header_columns(sheet)
            base_headers = {
                compact_text(cell.value): index + 1 for index, cell in enumerate(sheet[1])
            }
            equipment_keys = find_equipment_invoice_keys(workbook, invoice_file)
            period_invoice_total = ZERO
            period_equipment_total = ZERO

            for row_index in range(2, sheet.max_row + 1):
                invoice_date = parse_invoice_date(sheet.cell(row_index, columns["开票日期"]).value)
                seller = sheet.cell(row_index, columns["销方名称"]).value
                if (
                    invoice_date is None
                    or invoice_date.year != year
                    or invoice_date.month != period_month
                    or seller != config["seller_name"]
                ):
                    continue
                amount = invoice_to_money(sheet.cell(row_index, columns["金额"]).value)
                period_invoice_total += amount
                if is_equipment_invoice(sheet, row_index, base_headers, equipment_keys):
                    period_equipment_total += amount

            cumulative_invoice_total += period_invoice_total
            cumulative_equipment_total += period_equipment_total
            if period_month == month:
                monthly_invoice_total = period_invoice_total
                monthly_equipment_total = period_equipment_total
            workbook.close()

        overrides[profit_company] = {
            "monthly": monthly_invoice_total - monthly_equipment_total,
            "cumulative": cumulative_invoice_total - cumulative_equipment_total,
            "monthly_invoice_total": monthly_invoice_total,
            "monthly_equipment_total": monthly_equipment_total,
            "cumulative_invoice_total": cumulative_invoice_total,
            "cumulative_equipment_total": cumulative_equipment_total,
        }
    return overrides


def write_invoice_revenue_override_output(
    output_dir: Path,
    overrides: dict[str, dict[str, Decimal]],
    parsed: dict[str, dict[str, dict[str, Any]]],
) -> None:
    """输出科技、阳原开票营业收入取数口径，保留与利润表的差异。"""
    rows: list[dict[str, str]] = []
    for source_company, profit_company in INVOICE_REVENUE_COMPANY_MAP.items():
        values = overrides[profit_company]
        profit_values = parsed.get(profit_company, {}).get(
            "营业收入", {"monthly": ZERO, "cumulative": ZERO}
        )
        for scope, prefix, label in (
            ("本月", "monthly", "本月"),
            ("累计", "cumulative", "本年累计"),
        ):
            invoice_total = values[f"{prefix}_invoice_total"]
            equipment_total = values[f"{prefix}_equipment_total"]
            selected_revenue = values[prefix]
            profit_revenue = profit_values[prefix]
            rows.append(
                {
                    "主体": source_company,
                    "期间": label,
                    "利润表营业收入": format(profit_revenue, ".2f"),
                    "开票总额": format(invoice_total, ".2f"),
                    "设备发票金额": format(equipment_total, ".2f"),
                    "M3采用营业收入": format(selected_revenue, ".2f"),
                    "与利润表差异": format(selected_revenue - profit_revenue, ".2f"),
                    "说明": "M5将继续扣除M4内部销售，得到开票净营业收入",
                }
            )
    path = output_dir / "04_科技阳原营业收入开票口径.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def read_internal_sales(
    file_path: Path,
    source_company: str,
    year: int,
    month: int,
    group_company_names: frozenset[str] = GROUP_COMPANY_NAMES,
    buyer_company_names: frozenset[str] | None = None,
) -> list[InternalSaleRecord]:
    """读取一份开票文件；普通模式可兼容阳原文件错误的工作表维度信息。"""
    workbook = load_workbook(file_path, read_only=False, data_only=True)
    sheet = find_invoice_sheet(workbook, source_company)
    columns = header_columns(sheet)
    base_headers = {compact_text(cell.value): index + 1 for index, cell in enumerate(sheet[1])}
    equipment_keys = find_equipment_invoice_keys(workbook, file_path)
    config = SOURCE_CONFIG[source_company]
    buyer_company_names = group_company_names if buyer_company_names is None else buyer_company_names
    records: list[InternalSaleRecord] = []

    for row_index in range(2, sheet.max_row + 1):
        seller = sheet.cell(row_index, columns["销方名称"]).value
        buyer = sheet.cell(row_index, columns["购买方名称"]).value
        invoice_date = parse_invoice_date(sheet.cell(row_index, columns["开票日期"]).value)
        if invoice_date is None or invoice_date.year != year or invoice_date.month != month:
            continue
        # 销方与购买方相同属于异常开票记录，不能计入内部销售。
        if seller == config["seller_name"] and buyer == seller:
            raise ValueError(
                "发现销方与购买方相同的开票记录，无法继续计算内部销售："
                f"文件={file_path.name}，工作表={sheet.title}，行号={row_index}，"
                f"销方/购买方={seller}，开票日期={invoice_date.isoformat()}"
            )

        # 内部销售按固定集团主体名单判断，销方与购买方必须为不同主体。
        if (
            seller != config["seller_name"]
            or not is_group_company(seller, group_company_names)
            or not is_group_company(buyer, buyer_company_names)
        ):
            continue
        if is_equipment_invoice(sheet, row_index, base_headers, equipment_keys):
            continue
        amount = invoice_to_money(sheet.cell(row_index, columns["金额"]).value)
        status = "" if columns["发票状态"] == 0 else str(sheet.cell(row_index, columns["发票状态"]).value or "")
        records.append(
            InternalSaleRecord(
                period=f"{year}-{month:02d}",
                source_company=source_company,
                seller_name=seller,
                buyer_name=buyer,
                invoice_date=invoice_date.isoformat(),
                amount=amount,
                source_file=file_path.name,
                source_sheet=sheet.title,
                source_row=row_index,
                invoice_status=status,
            )
        )
    workbook.close()
    return records


def write_internal_outputs(
    output_dir: Path,
    cumulative_records: list[InternalSaleRecord],
    year: int,
    month: int,
    monthly_records: list[InternalSaleRecord] | None = None,
) -> None:
    """本月和累计可使用不同的集团主体范围。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    monthly_records = monthly_records if monthly_records is not None else [
        record for record in cumulative_records if record.period == f"{year}-{month:02d}"
    ]

    detail_path = output_dir / "01_内部销售明细.csv"
    with detail_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=(
                "期间", "销方主体", "销方名称", "购买方名称", "开票日期", "金额",
                "来源文件", "来源工作表", "来源行号", "发票状态",
            ),
        )
        writer.writeheader()
        for record in cumulative_records:
            writer.writerow(
                {
                    "期间": record.period,
                    "销方主体": record.source_company,
                    "销方名称": record.seller_name,
                    "购买方名称": record.buyer_name,
                    "开票日期": record.invoice_date,
                    "金额": format(record.amount, ".2f"),
                    "来源文件": record.source_file,
                    "来源工作表": record.source_sheet,
                    "来源行号": record.source_row,
                    "发票状态": record.invoice_status,
                }
            )

    def append_scope_summary(
        summary_rows: list[dict[str, str]], scope: str, records: list[InternalSaleRecord]
    ) -> Decimal:
        grouped: dict[tuple[str, str], Decimal] = defaultdict(lambda: ZERO)
        for record in records:
            grouped[(record.source_company, record.buyer_name)] += record.amount
        for (source_company, buyer_name), total in grouped.items():
            summary_rows.append(
                {
                    "范围": scope,
                    "销方主体": source_company,
                    "购买方名称": buyer_name,
                    "内部销售金额": format(total, ".2f"),
                }
            )
        total = sum((record.amount for record in records), ZERO)
        summary_rows.append(
            {
                "范围": scope,
                "销方主体": "合计",
                "购买方名称": "全部内部购买方",
                "内部销售金额": format(total, ".2f"),
            }
        )
        return total

    summary_rows: list[dict[str, str]] = []
    monthly_total = append_scope_summary(summary_rows, "本月", monthly_records)
    cumulative_total = append_scope_summary(summary_rows, "累计", cumulative_records)

    summary_path = output_dir / "02_内部销售汇总.csv"
    with summary_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=("范围", "销方主体", "购买方名称", "内部销售金额"))
        writer.writeheader()
        writer.writerows(summary_rows)

    report = [
        f"# {year}年{month}月内部销售计算结果",
        "",
        "## 汇总结果",
        "",
        "| 金额类型 | 内部销售金额（元） |",
        "| --- | ---: |",
        f"| {month}月本月内部销售 | {monthly_total:,.2f} |",
        f"| 1-{month}月累计内部销售 | {cumulative_total:,.2f} |",
        "",
        "## 计算口径",
        "",
        "- 本月和累计可使用不同的集团主体范围；美国分公司按季度纳入时，仅进入累计范围。",
        "- 内部销售：销方和购买方均属于对应范围的集团主体，且销方与购买方必须不同。",
        "- 如发现销方与购买方相同的记录，视为异常开票数据，脚本立即报错并停止计算。",
        "- 科技开票文件只保留销方为浙江长安仁恒科技股份有限公司的记录；阳原开票文件只保留销方为阳原县仁恒精细粘土有限责任公司的记录。",
        "- 只读取票级基础信息工作表和金额列；明细工作表不参与计算，避免重复累计。",
        "",
        f"明细文件：{detail_path.name}",
        f"汇总文件：{summary_path.name}",
    ]
    (output_dir / "03_内部销售结果说明.md").write_text("\n".join(report) + "\n", encoding="utf-8")


def run_m4(
    source_dir: Path,
    year: int,
    month: int,
    output_dir: Path,
    monthly_group_company_names: frozenset[str] = GROUP_COMPANY_NAMES,
    cumulative_group_company_names: frozenset[str] = GROUP_COMPANY_NAMES,
    quarterly_month_start: int | None = None,
    quarterly_buyer_company_names: frozenset[str] = frozenset(),
    quarterly_buyer_max_month: int | None = None,
) -> tuple[Path, Path]:
    """M4：本月和累计按各自有效主体范围生成内部销售。"""
    cumulative_records: list[InternalSaleRecord] = []
    monthly_records: list[InternalSaleRecord] = []
    for period_month in range(1, month + 1):
        for source_company in SOURCE_CONFIG:
            invoice_file = find_invoice_file(
                source_dir,
                source_company,
                year,
                period_month,
                required=source_company in REQUIRED_INVOICE_SOURCES,
            )
            if invoice_file is None:
                continue
            # 累计：美国/新疆购方数据只截止到最近季末月（如7月流程只算到6月）
            cumulative_buyer_names = cumulative_group_company_names
            if (
                quarterly_buyer_company_names
                and quarterly_buyer_max_month is not None
                and period_month > quarterly_buyer_max_month
            ):
                cumulative_buyer_names = cumulative_group_company_names - quarterly_buyer_company_names
            cumulative_records.extend(
                read_internal_sales(
                    invoice_file,
                    source_company,
                    year,
                    period_month,
                    cumulative_buyer_names,
                )
            )
            if period_month == month:
                monthly_records.extend(
                    read_internal_sales(
                        invoice_file,
                        source_company,
                        year,
                        period_month,
                        monthly_group_company_names,
                    )
                )
            elif quarterly_month_start is not None and period_month >= quarterly_month_start:
                monthly_records.extend(
                    read_internal_sales(
                        invoice_file,
                        source_company,
                        year,
                        period_month,
                        monthly_group_company_names,
                        quarterly_buyer_company_names,
                    )
                )
    write_internal_outputs(output_dir, cumulative_records, year, month, monthly_records)
    return output_dir / "02_内部销售汇总.csv", output_dir / "01_内部销售明细.csv"


def calculate_yangyuan_self_sales(
    source_dir: Path,
    year: int,
    month: int,
    monthly_group_company_names: frozenset[str] = GROUP_COMPANY_NAMES,
    cumulative_group_company_names: frozenset[str] = GROUP_COMPANY_NAMES,
) -> dict[str, dict[str, Decimal]]:
    """按阳原开票总额减内部销售；本月和累计可使用不同主体范围。"""
    config = SOURCE_CONFIG["阳原"]
    monthly_total = ZERO
    monthly_internal = ZERO
    cumulative_total = ZERO
    cumulative_internal = ZERO

    for period_month in range(1, month + 1):
        invoice_file = find_invoice_file(source_dir, "阳原", year, period_month)
        workbook = load_workbook(invoice_file, read_only=False, data_only=True)
        sheet = find_invoice_sheet(workbook, "阳原")
        columns = header_columns(sheet)
        period_total = ZERO
        period_cumulative_internal = ZERO
        period_monthly_internal = ZERO

        for row_index in range(2, sheet.max_row + 1):
            invoice_date = parse_invoice_date(sheet.cell(row_index, columns["开票日期"]).value)
            seller = sheet.cell(row_index, columns["销方名称"]).value
            buyer = sheet.cell(row_index, columns["购买方名称"]).value
            if (
                invoice_date is None
                or invoice_date.year != year
                or invoice_date.month != period_month
                or seller != config["seller_name"]
            ):
                continue
            amount = invoice_to_money(sheet.cell(row_index, columns["金额"]).value)
            period_total += amount
            if buyer == seller:
                raise ValueError(
                    "发现阳原销方与购买方相同的开票记录，无法计算阳原自销："
                    f"文件={invoice_file.name}，工作表={sheet.title}，行号={row_index}，"
                    f"销方/购买方={seller}，开票日期={invoice_date.isoformat()}"
                )
            if is_group_company(buyer, cumulative_group_company_names):
                period_cumulative_internal += amount
            if is_group_company(buyer, monthly_group_company_names):
                period_monthly_internal += amount

        cumulative_total += period_total
        cumulative_internal += period_cumulative_internal
        if period_month == month:
            monthly_total = period_total
            monthly_internal = period_monthly_internal
        workbook.close()

    return {
        "本月": {
            "阳原开票总额": monthly_total,
            "阳原内部销售": monthly_internal,
            "阳原自销": monthly_total - monthly_internal,
        },
        "累计": {
            "阳原开票总额": cumulative_total,
            "阳原内部销售": cumulative_internal,
            "阳原自销": cumulative_total - cumulative_internal,
        },
    }


def write_yangyuan_self_sales_output(
    output_dir: Path,
    current_values: dict[str, dict[str, Decimal]],
    prior_values: dict[str, dict[str, Decimal]],
    us_revenue: dict[str, Decimal] | None = None,
    prior_us_revenue: dict[str, Decimal] | None = None,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    us_revenue = us_revenue or {"本月": ZERO, "累计": ZERO}
    prior_us_revenue = prior_us_revenue or {"本月": ZERO, "累计": ZERO}
    rows = []
    for scope in ("累计", "本月"):
        rows.append(
            {
                "范围": scope,
                "本年阳原开票总额": format(current_values[scope]["阳原开票总额"], ".2f"),
                "本年阳原内部销售": format(current_values[scope]["阳原内部销售"], ".2f"),
                "本年阳原自销": format(current_values[scope]["阳原自销"], ".2f"),
                "去年同期阳原开票总额": format(prior_values[scope]["阳原开票总额"], ".2f"),
                "去年同期阳原内部销售": format(prior_values[scope]["阳原内部销售"], ".2f"),
                "去年同期阳原自销": format(prior_values[scope]["阳原自销"], ".2f"),
                "本年美国自销": format(us_revenue[scope], ".2f"),
                "去年同期美国自销": format(prior_us_revenue[scope], ".2f"),
            }
        )
    path = output_dir / "03_阳原自销计算结果.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    return path
