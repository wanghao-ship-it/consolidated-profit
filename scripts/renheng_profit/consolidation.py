"""M5 合并、营业成本倒挤、M6 指标及正式模板回填。"""
from __future__ import annotations

import csv
from decimal import Decimal
import json
from pathlib import Path
from typing import Any

import xlrd
import xlwt
from xlutils.copy import copy as copy_xls_workbook

from .common import (
    ELIMINATION_ITEMS,
    M6_METRIC_ORDER,
    M6_METRIC_ROWS,
    M6_NO_CHANGE_METRICS,
    OPERATING_COST_COMPONENT_WEIGHTS,
    REQUIRED_INTERNAL_COLUMNS,
    REQUIRED_PRE_COLUMNS,
    SUMMARY_CUMULATIVE_COLUMN,
    SUMMARY_MONTHLY_COLUMN,
    TEMPLATE_BLANK_ZERO_ITEMS,
    TEMPLATE_ITEM_ROWS,
    ZERO,
    elimination_json_default,
    elimination_to_money,
    metric_to_text,
)


def read_pre_elimination_summary(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if reader.fieldnames is None:
            raise ValueError(f"抵销前汇总文件为空：{path}")
        missing = REQUIRED_PRE_COLUMNS - set(reader.fieldnames)
        if missing:
            raise ValueError(f"抵销前汇总文件缺少列：{'、'.join(sorted(missing))}")
        rows = list(reader)
    if not any(row["项目"] == "营业收入" for row in rows):
        raise ValueError("抵销前汇总文件缺少“营业收入”项目")
    if not any(row["项目"] == "营业成本" for row in rows):
        raise ValueError("抵销前汇总文件缺少“营业成本”项目")
    return rows


def read_internal_sales_totals(path: Path) -> tuple[Decimal, Decimal]:
    """只读取内部销售汇总中的合计行，避免把销方-购方明细重复相加。"""
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if reader.fieldnames is None:
            raise ValueError(f"内部销售汇总文件为空：{path}")
        missing = REQUIRED_INTERNAL_COLUMNS - set(reader.fieldnames)
        if missing:
            raise ValueError(f"内部销售汇总文件缺少列：{'、'.join(sorted(missing))}")
        rows = list(reader)

    totals = {}
    for scope in ("本月", "累计"):
        matched = [row for row in rows if row["范围"] == scope and row["销方主体"] == "合计"]
        if len(matched) != 1:
            raise ValueError(f"内部销售汇总文件中“{scope}合计”应有且仅有一行，当前为{len(matched)}行")
        totals[scope] = elimination_to_money(matched[0]["内部销售金额"])
    return totals["本月"], totals["累计"]


def build_eliminated_rows(
    pre_rows: list[dict[str, str]],
    monthly_internal_sales: Decimal,
    cumulative_internal_sales: Decimal,
) -> list[dict[str, str]]:
    result = []
    for source_row in pre_rows:
        item = source_row["项目"]
        pre_monthly = elimination_to_money(source_row[SUMMARY_MONTHLY_COLUMN])
        pre_cumulative = elimination_to_money(source_row[SUMMARY_CUMULATIVE_COLUMN])
        monthly_elimination = monthly_internal_sales if item in ELIMINATION_ITEMS else ZERO
        cumulative_elimination = cumulative_internal_sales if item in ELIMINATION_ITEMS else ZERO
        final_monthly = pre_monthly - monthly_elimination
        if item in ELIMINATION_ITEMS:
            final_cumulative = pre_cumulative - cumulative_elimination
            explanation = "本月和累计营业收入分别扣除对应期间内部销售"
        else:
            final_cumulative = pre_cumulative
            explanation = "不参与内部销售抵销"
        result.append(
            {
                "项目": item,
                "抵销前本月": format(pre_monthly, ".2f"),
                "本月内部销售抵销": format(monthly_elimination, ".2f"),
                "抵销后本月": format(final_monthly, ".2f"),
                "抵销前累计": format(pre_cumulative, ".2f"),
                "累计内部销售抵销": format(cumulative_elimination, ".2f"),
                "抵销后累计": format(final_cumulative, ".2f"),
                "说明": explanation,
            }
        )
    return result


def recalculate_operating_cost_after_elimination(
    result_rows: list[dict[str, str]],
) -> list[dict[str, str]]:
    """按抵销后营业收入倒挤营业成本，抵销前营业成本保留用于溯源。"""
    rows_by_item = {row["项目"]: row for row in result_rows}
    required_items = {"营业收入", "营业成本", *OPERATING_COST_COMPONENT_WEIGHTS}
    missing = required_items - set(rows_by_item)
    if missing:
        raise ValueError(f"营业成本倒挤缺少科目：{'、'.join(sorted(missing))}")

    cost_row = rows_by_item["营业成本"]
    formula_text = (
        "营业成本 = 最终营业收入 -（税金及附加+销售费用+管理费用+研发费用+"
        "财务费用+投资收益+营业利润+公允价值变动收益）"
    )
    for field_name in ("抵销后本月", "抵销后累计"):
        revenue = Decimal(rows_by_item["营业收入"][field_name])
        deductions = sum(
            (
                Decimal(rows_by_item[item][field_name]) * weight
                for item, weight in OPERATING_COST_COMPONENT_WEIGHTS.items()
            ),
            ZERO,
        )
        cost_row[field_name] = format(revenue - deductions, ".2f")
    cost_row["说明"] = f"不直接抵销内部销售；抵销后按公式倒挤。{formula_text}"
    return result_rows


def write_elimination_outputs(
    output_dir: Path,
    result_rows: list[dict[str, str]],
    pre_summary_path: Path,
    internal_summary_path: Path,
    internal_detail_path: Path | None,
    monthly_internal_sales: Decimal,
    cumulative_internal_sales: Decimal,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    result_path = output_dir / "01_剔除内部销售合并利润表.csv"
    with result_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(result_rows[0]))
        writer.writeheader()
        writer.writerows(result_rows)

    manifest = {
        "抵销前六主体汇总表": str(pre_summary_path.resolve()),
        "内部销售汇总表": str(internal_summary_path.resolve()),
        "内部销售明细表": str(internal_detail_path.resolve()) if internal_detail_path else "未提供",
        "本月内部销售": monthly_internal_sales,
        "累计内部销售": cumulative_internal_sales,
        "抵销规则": {
            "营业收入本月": "抵销前本月金额 - 本月内部销售",
            "营业收入累计": "抵销前累计金额 - 累计内部销售",
            "营业成本": "按抵销后营业收入倒挤",
            "其他项目": "不调整",
        },
    }
    (output_dir / "02_抵销来源与规则.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2, default=elimination_json_default), encoding="utf-8"
    )

    row_by_item = {row["项目"]: row for row in result_rows}
    report = [
        "# 剔除内部销售后的合并利润表",
        "",
        "## 内部销售抵销金额",
        "",
        "| 期间 | 内部销售金额（元） |",
        "| --- | ---: |",
        f"| 本月 | {monthly_internal_sales:,.2f} |",
        f"| 累计 | {cumulative_internal_sales:,.2f} |",
        "",

        "## 抵销结果",
        "",
        "| 项目 | 抵销前本月 | 本月抵销 | 抵销后本月 | 抵销前累计 | 累计抵销 | 抵销后累计 |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for item in ("营业收入", "营业成本"):
        row = row_by_item[item]
        report.append(
            f"| {item} | {Decimal(row['抵销前本月']):,.2f} | {Decimal(row['本月内部销售抵销']):,.2f} | {Decimal(row['抵销后本月']):,.2f} | {Decimal(row['抵销前累计']):,.2f} | {Decimal(row['累计内部销售抵销']):,.2f} | {Decimal(row['抵销后累计']):,.2f} |"
        )
    report.extend(
        [
            "",
            "## 溯源资料",
            "",
            f"- M3抵销前六主体汇总表：{pre_summary_path}",
            f"- M4内部销售汇总表：{internal_summary_path}",
            f"- M4内部销售明细表：{internal_detail_path if internal_detail_path else '未提供'}",

            "说明：M3抵销前汇总表和M4内部销售明细均保持原文件不变；本脚本只生成新的抵销后结果和来源清单。",
        ]
    )
    (output_dir / "03_抵销结果说明.md").write_text("\n".join(report) + "\n", encoding="utf-8")


def run_m5(
    pre_summary_path: Path,
    internal_summary_path: Path,
    internal_detail_path: Path,
    output_dir: Path,
) -> Path:
    """本月和累计营业收入分别扣除对应期间内部销售。"""
    pre_rows = read_pre_elimination_summary(pre_summary_path)
    monthly_internal_sales, cumulative_internal_sales = read_internal_sales_totals(
        internal_summary_path
    )
    result_rows = build_eliminated_rows(
        pre_rows,
        monthly_internal_sales,
        cumulative_internal_sales,
    )
    recalculate_operating_cost_after_elimination(result_rows)
    write_elimination_outputs(
        output_dir,
        result_rows,
        pre_summary_path,
        internal_summary_path,
        internal_detail_path,
        monthly_internal_sales,
        cumulative_internal_sales,
    )
    return output_dir / "01_剔除内部销售合并利润表.csv"


def build_prior_year_rows(
    pre_rows: list[dict[str, str]],
    monthly_ledger_revenue: Decimal,
    cumulative_ledger_revenue: Decimal,
) -> list[dict[str, str]]:
    """去年同期营业收入直接取销售台账金额（不含税），不再抵销内部销售。"""
    result = []
    for source_row in pre_rows:
        item = source_row["项目"]
        pre_monthly = elimination_to_money(source_row[SUMMARY_MONTHLY_COLUMN])
        pre_cumulative = elimination_to_money(source_row[SUMMARY_CUMULATIVE_COLUMN])
        if item == "营业收入":
            monthly_elimination = ZERO
            cumulative_elimination = ZERO
            final_monthly = monthly_ledger_revenue
            final_cumulative = cumulative_ledger_revenue
            explanation = "去年同期营业收入直接取销售台账金额（不含税），不进行内部销售抵销"
        else:
            monthly_elimination = ZERO
            cumulative_elimination = ZERO
            final_monthly = pre_monthly
            final_cumulative = pre_cumulative
            explanation = "不参与内部销售抵销"
        result.append(
            {
                "项目": item,
                "抵销前本月": format(pre_monthly, ".2f"),
                "本月内部销售抵销": format(monthly_elimination, ".2f"),
                "抵销后本月": format(final_monthly, ".2f"),
                "抵销前累计": format(pre_cumulative, ".2f"),
                "累计内部销售抵销": format(cumulative_elimination, ".2f"),
                "上月最终累计": format(ZERO, ".2f"),
                "抵销后累计": format(final_cumulative, ".2f"),
                "说明": explanation,
            }
        )
    return result


def read_result_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def write_result_rows(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def calculate_m6_metrics(
    result_rows: list[dict[str, str]], value_field: str
) -> dict[str, Decimal | None]:
    """基于M5最终金额计算模板底部经营指标。"""
    rows_by_item = {row["项目"]: row for row in result_rows}
    required_items = {"营业收入", "营业成本", "净利润", "利润总额"}
    missing = required_items - set(rows_by_item)
    if missing:
        raise ValueError(f"M6指标计算缺少科目：{'、'.join(sorted(missing))}")

    revenue = Decimal(rows_by_item["营业收入"][value_field])
    cost = Decimal(rows_by_item["营业成本"][value_field])
    net_profit = Decimal(rows_by_item["净利润"][value_field])
    total_profit = Decimal(rows_by_item["利润总额"][value_field])
    gross_profit = revenue - cost

    def divide(numerator: Decimal, denominator: Decimal) -> Decimal | None:
        return None if denominator == ZERO else numerator / denominator

    return {
        "毛利率": divide(gross_profit, revenue),
        "毛利": gross_profit,
        "营业净利率": divide(net_profit, revenue),
        "税前利率": divide(total_profit, revenue),
        "主营业务成本率": divide(cost, revenue),
    }


def build_m6_output_rows(
    prior_rows: list[dict[str, str]], current_rows: list[dict[str, str]]
) -> list[dict[str, str]]:
    """生成本月和累计指标的同期对比结果。"""
    output_rows = []
    for scope, value_field in (("累计", "抵销后累计"), ("本月", "抵销后本月")):
        prior_metrics = calculate_m6_metrics(prior_rows, value_field)
        current_metrics = calculate_m6_metrics(current_rows, value_field)
        for metric in M6_METRIC_ORDER:
            prior_value = prior_metrics[metric]
            current_value = current_metrics[metric]
            change = (
                None
                if metric in M6_NO_CHANGE_METRICS or prior_value is None or current_value is None
                else current_value - prior_value
            )
            change_ratio = (
                None
                if change is None or prior_value is None or prior_value == ZERO
                else change / prior_value
            )
            is_amount = metric == "毛利"
            output_rows.append(
                {
                    "范围": scope,
                    "指标": metric,
                    "去年同期": metric_to_text(prior_value, is_amount),
                    "本年": metric_to_text(current_value, is_amount),
                    "增减金额": metric_to_text(change, is_amount),
                    "增减比率": metric_to_text(change_ratio),
                }
            )
    return output_rows


def write_m6_outputs(output_dir: Path, rows: list[dict[str, str]]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "01_经营指标计算结果.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    return path


def write_template_cell(
    read_book,
    write_book,
    sheet_index: int,
    row_index: int,
    column_index: int,
    value: Any,
) -> None:
    """使用复制后单元格的原样式写入新值。"""
    write_sheet = write_book.get_sheet(sheet_index)
    style_collection = write_book._Workbook__styles
    row_object = write_sheet._Worksheet__rows.get(row_index)
    if row_object is None:
        write_sheet.write(row_index, column_index, value)
        return
    existing_cell = row_object._Row__cells.get(column_index)
    if existing_cell is None:
        # 模板底部部分指标行没有预置全部C-F单元格，使用同一行已有指标单元格样式。
        existing_cell = next(
            (row_object._Row__cells.get(index) for index in (2, 3, 4, 5) if row_object._Row__cells.get(index) is not None),
            None,
        )
    if existing_cell is None:
        write_sheet.write(row_index, column_index, value)
        return
    font_index, number_format_index, alignment, borders, pattern, protection = (
        style_collection._xf_x2id[existing_cell.xf_idx]
    )
    number_format = next(
        (
            format_text
            for format_text, format_index in style_collection._num_formats.items()
            if format_index == number_format_index and format_text is not None
        ),
        "General",
    )
    style = xlwt.XFStyle()
    style.font = style_collection._font_x2id[font_index]
    style.num_format_str = number_format
    style.alignment = alignment
    style.borders = borders
    style.pattern = pattern
    style.protection = protection
    write_sheet.write(row_index, column_index, value, style)


def write_final_template(
    template_path: Path,
    output_path: Path,
    current_rows: list[dict[str, str]],
    prior_rows: list[dict[str, str]],
    current_yangyuan_self_sales: dict[str, dict[str, Decimal]],
    prior_yangyuan_self_sales: dict[str, dict[str, Decimal]],
    year: int,
    month: int,
    us_revenue: dict[str, Decimal] | None = None,
    us_prior_revenue: dict[str, Decimal] | None = None,
) -> None:
    """复制.xls模板，只回填两张利润表的核心表格区域。"""
    read_book = xlrd.open_workbook(template_path, formatting_info=True)
    us_revenue = us_revenue or {"本月": ZERO, "累计": ZERO}
    us_prior_revenue = us_prior_revenue or {"本月": ZERO, "累计": ZERO}
    write_book = copy_xls_workbook(read_book)
    style_collection = write_book._Workbook__styles
    invalid_format_index = style_collection._num_formats.pop(None, None)
    if invalid_format_index is not None:
        style_collection._num_formats["General "] = invalid_format_index
    current_by_item = {row["项目"]: row for row in current_rows}
    prior_by_item = {row["项目"]: row for row in prior_rows}
    prior_year = year - 1

    sheet_specs = (
        (0, "抵销后累计", "去年同期", "本年合计"),
        (1, "抵销后本月", f"{str(prior_year)[-2:]}年{month}月", f"{str(year)[-2:]}年{month}月"),
    )
    for sheet_index, value_field, prior_header, current_header in sheet_specs:
        write_template_cell(read_book, write_book, sheet_index, 3, 2, prior_header)
        write_template_cell(read_book, write_book, sheet_index, 3, 3, current_header)
        for item, row_index in TEMPLATE_ITEM_ROWS.items():
            prior_value = Decimal(prior_by_item[item][value_field])
            current_value = Decimal(current_by_item[item][value_field])
            if item in TEMPLATE_BLANK_ZERO_ITEMS and prior_value == ZERO and current_value == ZERO:
                values: tuple[Any, Any, Any, Any] = ("", "", "", "")
            else:
                difference = current_value - prior_value
                ratio: Any = "" if prior_value == ZERO else float(difference / prior_value)
                values = (
                    float(prior_value),
                    float(current_value),
                    float(difference),
                    ratio,
                )
            for offset, value in enumerate(values, start=2):
                write_template_cell(
                    read_book, write_book, sheet_index, row_index, offset, value
                )

        prior_metrics = calculate_m6_metrics(prior_rows, value_field)
        current_metrics = calculate_m6_metrics(current_rows, value_field)
        for metric, row_index in M6_METRIC_ROWS.items():
            prior_value = prior_metrics[metric]
            current_value = current_metrics[metric]
            change = (
                None
                if metric in M6_NO_CHANGE_METRICS or prior_value is None or current_value is None
                else current_value - prior_value
            )
            change_ratio = (
                None
                if change is None or prior_value is None or prior_value == ZERO
                else change / prior_value
            )
            is_amount = metric == "毛利"
            metric_values: tuple[Any, Any, Any, Any] = (
                "" if prior_value is None else float(prior_value),
                "" if current_value is None else float(current_value),
                "" if change is None else float(change),
                "" if change_ratio is None else float(change_ratio),
            )
            for offset, value in enumerate(metric_values, start=2):
                write_template_cell(
                    read_book, write_book, sheet_index, row_index, offset, value
                )

        # 清除模板底部遗留说明和空白行，不清除指标标签及C-F回填结果。
        metric_row_indexes = set(M6_METRIC_ROWS.values())
        write_sheet = write_book.get_sheet(sheet_index)
        for row_index in range(28, 38):
            row_object = write_sheet._Worksheet__rows.get(row_index)
            if row_object is None:
                continue
            for column_index in range(8):
                keep_metric_value = row_index in metric_row_indexes and column_index <= 5
                if not keep_metric_value and column_index in row_object._Row__cells:
                    write_template_cell(
                        read_book,
                        write_book,
                        sheet_index,
                        row_index,
                        column_index,
                        "",
                    )

        # 指标说明文字（参照标准报表版式）：
        # E29 毛利率增减（个百分点文字，F29 保持增减比率数值）
        prior_gross_rate = prior_metrics["毛利率"]
        current_gross_rate = current_metrics["毛利率"]
        if prior_gross_rate is not None and current_gross_rate is not None:
            rate_change_pp = (current_gross_rate - prior_gross_rate) * 100
            rate_text = (
                f"增长了{abs(rate_change_pp):.2f}个百分点"
                if rate_change_pp >= 0
                else f"下降了{abs(rate_change_pp):.2f}个百分点"
            )
            write_template_cell(read_book, write_book, sheet_index, 28, 4, rate_text)
        # E31 毛利增减（万元文字）
        prior_gross = prior_metrics["毛利"]
        current_gross = current_metrics["毛利"]
        if prior_gross is not None and current_gross is not None:
            gross_change_wan = round((current_gross - prior_gross) / 10000, 1)
            gross_text = (
                f"比去年同期增加毛利{abs(gross_change_wan):.1f}万"
                if gross_change_wan >= 0
                else f"比去年同期减少毛利{abs(gross_change_wan):.1f}万"
            )
            write_template_cell(read_book, write_book, sheet_index, 30, 4, gross_text)

        self_sale_scope = "累计" if sheet_index == 0 else "本月"
        scope_label = "" if sheet_index == 0 else "本月"
        write_template_cell(
            read_book, write_book, sheet_index, 4, 6,
            f"{str(year)[-2:]}年{scope_label}其中阳原自销",
        )
        write_template_cell(
            read_book, write_book, sheet_index, 4, 7,
            float(current_yangyuan_self_sales[self_sale_scope]["阳原自销"]),
        )
        write_template_cell(
            read_book, write_book, sheet_index, 5, 6,
            f"{str(prior_year)[-2:]}年{scope_label}其中阳原自销",
        )
        write_template_cell(
            read_book, write_book, sheet_index, 5, 7,
            float(prior_yangyuan_self_sales[self_sale_scope]["阳原自销"]),
        )
        # 美国自销 = 美国利润表营业收入（本月/累计各取各的）；25年无数据填0
        write_template_cell(
            read_book, write_book, sheet_index, 6, 6,
            f"{str(year)[-2:]}年{scope_label}其中美国自销",
        )
        write_template_cell(
            read_book, write_book, sheet_index, 6, 7,
            float(us_revenue[self_sale_scope]),
        )
        write_template_cell(
            read_book, write_book, sheet_index, 7, 6,
            f"{str(prior_year)[-2:]}年{scope_label}其中美国自销",
        )
        write_template_cell(
            read_book, write_book, sheet_index, 7, 7,
            float(us_prior_revenue[self_sale_scope]),
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_book.save(str(output_path))
