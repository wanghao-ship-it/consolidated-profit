"""销售台账读取、M7 校验、主体拆分和无票销售归因。"""
from __future__ import annotations

import csv
from collections import defaultdict
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import re
from typing import Any

from .common import (
    GROUP_COMPANIES,
    GROUP_COMPANY_NAMES,
    LEDGER_BUYER_ALIASES,
    LEDGER_HEADER_ALIASES,
    LEDGER_INVOICE_NUMBER_ALIASES,
    US_COMPANY,
    ZERO,
    LedgerDataset,
    compact_text,
    invoice_to_money,
    is_group_company,
    to_number,
)
from .consolidation import elimination_to_money, read_pre_elimination_summary
from .income_statements import read_sheet_rows


def find_sales_ledger_file(source_dir: Path, year: int) -> Path:
    matches = [
        path
        for path in source_dir.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xls", ".xlsx"}
        and "销售台账" in compact_text(path.name)
        and str(year) in compact_text(path.name)
        and not path.name.startswith("~$")
    ]
    if len(matches) != 1:
        detail = "未找到" if not matches else f"找到{len(matches)}份"
        raise FileNotFoundError(
            f"{year}年销售台账{detail}；请保证目录中只有一份有效销售台账"
        )
    return matches[0]


def read_ledger_dataset(source_dir: Path, year: int, month: int) -> LedgerDataset:
    """强校验销售台账并保留逐行证据；票号列缺失是审核项而非解析失败。"""
    ledger_file = find_sales_ledger_file(source_dir, year)
    candidates: list[tuple[str, int, dict[str, int], list[list[Any]]]] = []
    for sheet_name, rows in read_sheet_rows(ledger_file):
        for header_row_index, header_row in enumerate(rows[:20]):
            headers = [compact_text(value) for value in header_row]
            columns: dict[str, int] = {}
            for field_name, aliases in LEDGER_HEADER_ALIASES.items():
                column = next((headers.index(alias) for alias in aliases if alias in headers), None)
                if column is None:
                    break
                columns[field_name] = column
            if len(columns) != len(LEDGER_HEADER_ALIASES):
                continue
            buyer_column = next((headers.index(alias) for alias in LEDGER_BUYER_ALIASES if alias in headers), None)
            if buyer_column is None:
                raise ValueError(f"{ledger_file.name}的销售台账缺少购买方名称列")
            columns["购买方名称"] = buyer_column
            invoice_column = next(
                (headers.index(alias) for alias in LEDGER_INVOICE_NUMBER_ALIASES if alias in headers),
                None,
            )
            if invoice_column is not None:
                columns["数电票号码"] = invoice_column
            candidates.append((sheet_name, header_row_index, columns, rows))
    if len(candidates) != 1:
        detail = "未找到" if not candidates else f"找到{len(candidates)}个"
        raise ValueError(
            f"{ledger_file.name}中{detail}销售台账明细表；必须唯一包含月份、销方名称、购买方名称、金额（不含税）"
        )
    sheet_name, header_row_index, columns, raw_rows = candidates[0]
    details: list[dict[str, Any]] = []
    required_max = max(columns[name] for name in LEDGER_HEADER_ALIASES)
    for source_row, row in enumerate(raw_rows[header_row_index + 1 :], start=header_row_index + 2):
        if required_max >= len(row):
            continue
        month_match = re.fullmatch(r"(\d{1,2})月?", compact_text(row[columns["月份"]]))
        if month_match is None:
            continue
        period_month = int(month_match.group(1))
        if not 1 <= period_month <= month:
            continue
        raw_amount = row[columns["金额（不含税）"]]
        amount = to_number(raw_amount)
        if amount is None:
            raise ValueError(
                f"{ledger_file.name}的{sheet_name}第{source_row}行金额（不含税）为空或无法解析"
            )
        invoice_number = (
            compact_text(row[columns["数电票号码"]])
            if "数电票号码" in columns and columns["数电票号码"] < len(row)
            else ""
        )
        details.append(
            {
                "月份": period_month,
                "销方": compact_text(row[columns["销方名称"]]),
                "购方": compact_text(row[columns["购买方名称"]]) if columns["购买方名称"] < len(row) else "",
                "票号": invoice_number,
                "金额": amount,
                "是否无票销售": invoice_number == "无票销售",
                "源文件": str(ledger_file),
                "源工作表": sheet_name,
                "源行号": source_row,
            }
        )
    return LedgerDataset(
        file=ledger_file,
        sheet=sheet_name,
        header_row=header_row_index + 1,
        has_invoice_number="数电票号码" in columns,
        rows=details,
    )


def read_prior_ledger_us_internal(
    ledger_path: Path, month: int, cumulative_max_month: int, monthly_min_month: int | None = None
) -> dict[str, Decimal]:
    """从去年销售台账读取科技→美国金额（累计=1~cumulative_max_month；本月=monthly_min_month~month，None 则本月为0）。"""
    ledger_file = Path(ledger_path)
    candidates: list[tuple[str, int, dict[str, int], list[list[Any]]]] = []
    for sheet_name, rows in read_sheet_rows(ledger_file):
        for header_row_index, header_row in enumerate(rows[:20]):
            compact_headers = [compact_text(value) for value in header_row]
            columns: dict[str, int] = {}
            for field_name, aliases in LEDGER_HEADER_ALIASES.items():
                column = next(
                    (compact_headers.index(a) for a in aliases if a in compact_headers),
                    None,
                )
                if column is None:
                    break
                columns[field_name] = column
            if len(columns) == len(LEDGER_HEADER_ALIASES):
                if "购买方名称" in compact_headers:
                    columns["购买方名称"] = compact_headers.index("购买方名称")
                candidates.append((sheet_name, header_row_index, columns, rows))
    if len(candidates) != 1:
        detail = "未找到" if not candidates else f"找到{len(candidates)}个"
        raise ValueError(f"{ledger_file.name}中{detail}销售台账明细表")
    _, header_row_index, columns, rows = candidates[0]
    tech_name = GROUP_COMPANIES["科技"]
    us_names = frozenset({GROUP_COMPANIES[US_COMPANY]})
    monthly = ZERO
    cumulative = ZERO
    for row in rows[header_row_index + 1:]:
        if max(columns.values()) >= len(row):
            continue
        if compact_text(row[columns["销方名称"]]) != tech_name:
            continue
        buyer = compact_text(row[columns.get("购买方名称", 0)]) if "购买方名称" in columns else ""
        if not is_group_company(buyer, us_names):
            continue
        month_match = re.fullmatch(r"(\d{1,2})月", compact_text(row[columns["月份"]]))
        if month_match is None:
            continue
        period_month = int(month_match.group(1))
        amount = invoice_to_money(row[columns["金额（不含税）"]])
        if 1 <= period_month <= cumulative_max_month:
            cumulative += amount
        if monthly_min_month is not None and monthly_min_month <= period_month <= month:
            monthly += amount
    return {"本月": monthly, "累计": cumulative}


def read_prior_sales_ledger(ledger_path: Path, month: int) -> dict[str, dict[str, Decimal]]:
    """从去年销售台账直接计算阳原自销。

    销售台账为对外销售口径（不含内部销售、含无票销售），阳原销方金额即阳原自销。
    返回结构与 calculate_yangyuan_self_sales 一致；仅"阳原自销"有值，
    "阳原开票总额""阳原内部销售"置0（台账无法区分）。
    """
    ledger_file = Path(ledger_path)
    if not ledger_file.is_file():
        raise FileNotFoundError(f"去年销售台账文件不存在：{ledger_file}")
    candidates: list[tuple[str, int, dict[str, int], list[list[Any]]]] = []
    for sheet_name, rows in read_sheet_rows(ledger_file):
        for header_row_index, header_row in enumerate(rows[:20]):
            compact_headers = [compact_text(value) for value in header_row]
            columns: dict[str, int] = {}
            for field_name, aliases in LEDGER_HEADER_ALIASES.items():
                column = next(
                    (
                        compact_headers.index(alias)
                        for alias in aliases
                        if alias in compact_headers
                    ),
                    None,
                )
                if column is None:
                    break
                columns[field_name] = column
            if len(columns) == len(LEDGER_HEADER_ALIASES):
                candidates.append((sheet_name, header_row_index, columns, rows))
    if len(candidates) != 1:
        detail = "未找到" if not candidates else f"找到{len(candidates)}个"
        raise ValueError(
            f"{ledger_file.name}中{detail}销售台账明细表；需要包含“月份、销方名称、金额（不含税）”列"
        )
    sheet_name, header_row_index, columns, rows = candidates[0]
    yangyuan_name = GROUP_COMPANIES["阳原"]
    monthly = ZERO
    cumulative = ZERO
    for row in rows[header_row_index + 1:]:
        if max(columns.values()) >= len(row):
            continue
        if compact_text(row[columns["销方名称"]]) != yangyuan_name:
            continue
        month_match = re.fullmatch(r"(\d{1,2})月", compact_text(row[columns["月份"]]))
        if month_match is None:
            continue
        period_month = int(month_match.group(1))
        if not 1 <= period_month <= month:
            continue
        amount = invoice_to_money(row[columns["金额（不含税）"]])
        cumulative += amount
        if period_month == month:
            monthly += amount
    return {
        "本月": {"阳原开票总额": ZERO, "阳原内部销售": ZERO, "阳原自销": monthly},
        "累计": {"阳原开票总额": ZERO, "阳原内部销售": ZERO, "阳原自销": cumulative},
    }


def read_final_revenue_values(result_rows: list[dict[str, str]]) -> dict[str, Decimal]:
    revenue_row = next((row for row in result_rows if row["项目"] == "营业收入"), None)
    if revenue_row is None:
        raise ValueError("M5抵销后结果缺少“营业收入”项目，无法执行M7校验")
    return {
        "本月": elimination_to_money(revenue_row["抵销后本月"]),
        "累计": elimination_to_money(revenue_row["抵销后累计"]),
    }


def read_us_xinjiang_revenue(pre_summary_path: Path) -> dict[str, Decimal]:
    """从M3抵销前汇总读取美国/新疆营业收入（本月/累计），用于M7国内主体校验调整。"""
    rows = read_pre_elimination_summary(pre_summary_path)
    revenue_row = next(row for row in rows if row["项目"] == "营业收入")

    def company_value(company: str, scope_label: str) -> Decimal:
        return elimination_to_money(revenue_row.get(f"{company}_{scope_label}", "0"))

    return {
        "本月": company_value(US_COMPANY, "本月") + company_value("新疆公司", "本月"),
        "累计": company_value(US_COMPANY, "累计") + company_value("新疆公司", "累计"),
    }


def read_us_buyer_internal(
    detail_path: Path, max_month: int, target_month: int | None = None
) -> dict[str, Decimal]:
    """统计内部销售明细中购方=美国的金额（月份<=max_month），用于M7台账扣减。"""
    us_buyer_names = frozenset({GROUP_COMPANIES[US_COMPANY]})
    monthly = ZERO
    cumulative = ZERO
    with open(detail_path, encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            period = row.get("期间", "")
            if len(period) < 7:
                continue
            period_month = int(period[5:7])
            if period_month > max_month:
                continue
            buyer = row.get("购买方名称", "")
            if not is_group_company(buyer, us_buyer_names):
                continue
            amount = invoice_to_money(row.get("金额", "0"))
            cumulative += amount
            if target_month is not None and period_month == target_month:
                monthly += amount
    return {"本月": monthly, "累计": cumulative}


def build_m7_validation_rows(
    current_result_rows: list[dict[str, str]],
    prior_result_rows: list[dict[str, str]],
    current_ledger: dict[str, Any],
    prior_ledger: dict[str, Any] | None,
    current_revenue_adjust: dict[str, Decimal] | None = None,
    current_ledger_adjust: dict[str, Decimal] | None = None,
    prior_revenue_adjust: dict[str, Decimal] | None = None,
    prior_ledger_adjust: dict[str, Decimal] | None = None,
) -> list[dict[str, str]]:
    """M7校验；美国/新疆纳入后，营业收入扣除美新收入、台账扣除美国购方内销，仅比较国内主体口径。"""
    rows = []
    current_revenue = read_final_revenue_values(current_result_rows)
    prior_revenue = read_final_revenue_values(prior_result_rows)
    for scope in ("累计", "本月"):
        for label, final_revenue, ledger, rev_adjust, led_adjust in (
            ("本年", current_revenue, current_ledger, current_revenue_adjust, current_ledger_adjust),
            ("去年同期", prior_revenue, prior_ledger, prior_revenue_adjust, prior_ledger_adjust),
        ):
            if label == "去年同期" and ledger is None:
                rows.append(
                    {
                        "年度类型": label,
                        "范围": scope,
                        "M5抵销后营业收入": format(final_revenue[scope], ".2f"),
                        "销售台账金额（不含税）": "",
                        "差异": "",
                        "校验结果": "不适用（去年最终合并报表直取）",
                        "校验说明": "去年同期数据直接复制去年最终合并报表，未读取销售台账",
                        "销售台账文件": "",
                        "工作表": "",
                        "台账表头行": "",
                    }
                )
                continue

            ledger_value = ledger[scope]
            final_value = final_revenue[scope]
            if rev_adjust:
                final_value -= rev_adjust[scope]
            if led_adjust:
                ledger_value -= led_adjust[scope]
            difference = final_value - ledger_value
            result = "一致" if abs(difference) == ZERO else "存在差异"
            note = (
                f"{label}M5抵销后营业收入与销售台账独立校验（仅国内主体口径，已扣除美国/新疆）"
            )
            rows.append(
                {
                    "年度类型": label,
                    "范围": scope,
                    "M5抵销后营业收入": format(final_value, ".2f"),
                    "销售台账金额（不含税）": format(ledger_value, ".2f"),
                    "差异": format(difference, ".2f"),
                    "校验结果": result,
                    "校验说明": note,
                    "销售台账文件": ledger["file"],
                    "工作表": ledger["sheet"],
                    "台账表头行": str(ledger["header_row"]),
                }
            )
    return rows


def build_current_subject_m7_breakdown(
    current_pre_summary: Path,
    current_internal_summary: Path,
    current_ledger: dict[str, Any],
    m7_rows: list[dict[str, str]],
) -> list[dict[str, str]]:
    """按销方主体直接比较M5收入贡献与销售台账，并勾稽到本年集团M7差异。

    科技、阳原是本年采用开票口径覆盖利润表收入的两个主体，因此固定逐项披露；
    其他主体或口径调整造成的剩余金额不得强行分配，单列为“其他主体/未定位”。
    """
    pre_rows = read_pre_elimination_summary(current_pre_summary)
    revenue_row = next((row for row in pre_rows if row["项目"] == "营业收入"), None)
    if revenue_row is None:
        raise ValueError("M3抵销前汇总缺少营业收入，无法生成主体M7差异拆分")

    internal_by_scope_and_seller: dict[tuple[str, str], Decimal] = defaultdict(lambda: ZERO)
    us_internal_by_scope_and_seller: dict[tuple[str, str], Decimal] = defaultdict(lambda: ZERO)
    with current_internal_summary.open(encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            scope = row.get("范围", "")
            seller = row.get("销方主体", "")
            if scope not in {"本月", "累计"} or seller == "合计":
                continue
            amount = elimination_to_money(row.get("内部销售金额", "0"))
            internal_by_scope_and_seller[(scope, seller)] += amount
            if is_group_company(row.get("购买方名称", ""), frozenset({GROUP_COMPANIES[US_COMPANY]})):
                us_internal_by_scope_and_seller[(scope, seller)] += amount

    group_difference = {
        row["范围"]: elimination_to_money(row["差异"])
        for row in m7_rows
        if row["年度类型"] == "本年"
    }
    subject_config = (
        ("科技", "仁恒科技", GROUP_COMPANIES["科技"]),
        ("阳原", "阳原仁恒", GROUP_COMPANIES["阳原"]),
    )
    rows: list[dict[str, str]] = []
    subject_totals: dict[str, Decimal] = {"本月": ZERO, "累计": ZERO}
    for scope, m3_scope in (("本月", "本月"), ("累计", "累计")):
        if scope not in group_difference:
            raise ValueError(f"M7缺少本年{scope}集团差异，无法生成主体差异拆分")
        ledger_by_company = current_ledger[f"{scope}_按主体"]
        for seller_key, profit_company, legal_name in subject_config:
            m3_revenue = elimination_to_money(revenue_row.get(f"{profit_company}_{m3_scope}", "0"))
            internal_sales = internal_by_scope_and_seller[(scope, seller_key)]
            subject_m5_revenue = m3_revenue - internal_sales
            ledger_revenue = elimination_to_money(ledger_by_company.get(legal_name, ZERO))
            # 集团M7在台账端扣除美国购方内部销售；主体拆分必须采用同一口径。
            ledger_comparable = ledger_revenue - us_internal_by_scope_and_seller[(scope, seller_key)]
            difference = subject_m5_revenue - ledger_comparable
            subject_totals[scope] += difference
            rows.append(
                {
                    "范围": scope,
                    "主体": profit_company,
                    "主体差异": format(difference, ".2f"),
                    "对集团差异的影响": format(difference, ".2f"),
                }
            )

        residual = group_difference[scope] - subject_totals[scope]
        if residual != ZERO:
            rows.append(
                {
                    "范围": scope,
                    "主体": "其他主体/未定位",
                    "主体差异": format(residual, ".2f"),
                    "对集团差异的影响": format(residual, ".2f"),
                }
            )

    rows.append(
        {
            "范围": "累计合计",
            "主体": "—",
            "主体差异": "—",
            "对集团差异的影响": format(group_difference["累计"], ".2f"),
        }
    )
    return rows


def write_subject_m7_breakdown(output_dir: Path, rows: list[dict[str, str]]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "05_主体销售台账差异拆分.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(
            file, fieldnames=("范围", "主体", "主体差异", "对集团差异的影响")
        )
        writer.writeheader()
        writer.writerows(rows)
    return path


def write_m7_validation_outputs(output_dir: Path, rows: list[dict[str, str]]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / "01_销售台账营业收入校验.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)

    report_lines = [
        "# 销售台账与抵销后营业收入校验",
        "",
        "校验公式：差异 = M5抵销后营业收入 - 销售台账金额（不含税）。",
        "销售台账仅用于M7校验，不参与M4内部销售计算、M5营业收入抵销、营业成本倒挤或模板金额回填。",
        "校验结果：差异绝对值严格等于0.00时为“一致”；差异绝对值不为0.00时为“存在差异”。",
        "",
        "| 年度类型 | 范围 | M5抵销后营业收入 | 销售台账金额（不含税） | 差异 | 结果 | 说明 |",
        "| --- | --- | ---: | ---: | ---: | --- | --- |",
    ]
    for row in rows:
        report_lines.append(
            f"| {row['年度类型']} | {row['范围']} | {row['M5抵销后营业收入']} | "
            f"{row['销售台账金额（不含税）']} | {row['差异']} | {row['校验结果']} | {row['校验说明']} |"
        )
    (output_dir / "02_销售台账校验说明.md").write_text(
        "\n".join(report_lines) + "\n", encoding="utf-8"
    )
    return csv_path


def ledger_summary(
    dataset: LedgerDataset,
    month: int,
    seller_names: frozenset[str] | None = None,
) -> dict[str, Any]:
    monthly_by_company = {name: ZERO for name in GROUP_COMPANIES.values()}
    cumulative_by_company = {name: ZERO for name in GROUP_COMPANIES.values()}
    for row in dataset.rows:
        seller = row["销方"]
        if seller not in GROUP_COMPANY_NAMES:
            continue
        if seller_names is not None and seller not in seller_names:
            continue
        cumulative_by_company[seller] += row["金额"]
        if row["月份"] == month:
            monthly_by_company[seller] += row["金额"]
    return {
        "file": str(dataset.file),
        "sheet": dataset.sheet,
        "header_row": dataset.header_row,
        "本月": sum(monthly_by_company.values(), ZERO),
        "累计": sum(cumulative_by_company.values(), ZERO),
        "本月_按主体": monthly_by_company,
        "累计_按主体": cumulative_by_company,
    }


def unbilled_evidence(
    dataset: LedgerDataset,
    year: int,
    target_month: int,
    scope: str,
    exclude_us_seller: bool = False,
) -> list[dict[str, Any]]:
    return [
        {
            "年度": year,
            "范围": scope,
            **row,
        }
        for row in dataset.rows
        if row["是否无票销售"]
        and row["销方"] in GROUP_COMPANY_NAMES
        and not (exclude_us_seller and row["销方"] == GROUP_COMPANIES[US_COMPANY])
        and (row["月份"] == target_month if scope == "本月" else row["月份"] <= target_month)
    ]


def enrich_m7_rows_with_unbilled(
    rows: list[dict[str, str]],
    current_dataset: LedgerDataset,
    prior_dataset: LedgerDataset,
    year: int,
    month: int,
) -> tuple[list[dict[str, str]], dict[str, list[dict[str, Any]]]]:
    evidence_by_diff: dict[str, list[dict[str, Any]]] = {}
    for index, row in enumerate(rows, start=1):
        diff_id = f"M7-{index:02d}"
        row["差异ID"] = diff_id
        dataset = current_dataset if row["年度类型"] == "本年" else prior_dataset
        is_prior = row["年度类型"] == "去年同期"
        evidence = unbilled_evidence(
            dataset,
            year - 1 if is_prior else year,
            month,
            row["范围"],
            exclude_us_seller=is_prior,
        )
        unbilled_total = sum((item["金额"] for item in evidence), ZERO)
        difference = elimination_to_money(row["差异"])
        residual = difference + unbilled_total
        if residual == ZERO and unbilled_total != ZERO:
            classification = "无票销售完全解释"
        elif unbilled_total != ZERO and abs(residual) < abs(difference):
            classification = "无票销售部分解释"
        else:
            classification = "无票销售未解释"
        row["无票销售金额"] = format(unbilled_total, ".2f")
        row["无票后剩余差异"] = format(residual, ".2f")
        row["无票归因"] = classification
        evidence_by_diff[diff_id] = evidence
    return rows, evidence_by_diff


def diff_snapshot_hash(rows: list[dict[str, str]]) -> str:
    payload = [
        {
            "差异ID": row["差异ID"],
            "年度类型": row["年度类型"],
            "范围": row["范围"],
            "差异": row["差异"],
            "无票销售金额": row["无票销售金额"],
            "无票后剩余差异": row["无票后剩余差异"],
        }
        for row in rows
    ]
    return hashlib.sha256(
        json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()


def nonzero_m7_keys(rows: list[dict[str, str]]) -> set[tuple[str, str]]:
    return {
        (row["年度类型"], row["范围"])
        for row in rows
        if elimination_to_money(row["差异"]) != ZERO
    }
