import csv
import importlib.util
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook


SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import renheng_profit as WORKFLOW

GRAPH_SPEC = importlib.util.spec_from_file_location(
    "langgraph_profit_workflow_for_subject_tests",
    SCRIPTS / "langgraph_profit_workflow.py",
)
GRAPH_WORKFLOW = importlib.util.module_from_spec(GRAPH_SPEC)
sys.modules[GRAPH_SPEC.name] = GRAPH_WORKFLOW
GRAPH_SPEC.loader.exec_module(GRAPH_WORKFLOW)


class SubjectM7BreakdownTests(unittest.TestCase):
    def make_inputs(self, directory: Path):
        pre_summary = directory / "pre.csv"
        with pre_summary.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(
                file,
                fieldnames=(
                    "项目",
                    "六主体汇总_本月",
                    "六主体汇总_累计",
                    "仁恒科技_本月",
                    "仁恒科技_累计",
                    "阳原仁恒_本月",
                    "阳原仁恒_累计",
                ),
            )
            writer.writeheader()
            writer.writerow(
                {
                    "项目": "营业收入",
                    "六主体汇总_本月": "0.00",
                    "六主体汇总_累计": "0.00",
                    "仁恒科技_本月": "8381608.67",
                    "仁恒科技_累计": "24661034.29",
                    "阳原仁恒_本月": "7768537.66",
                    "阳原仁恒_累计": "17442131.80",
                }
            )
            writer.writerow(
                {
                    "项目": "营业成本",
                    "六主体汇总_本月": "0.00",
                    "六主体汇总_累计": "0.00",
                    "仁恒科技_本月": "0.00",
                    "仁恒科技_累计": "0.00",
                    "阳原仁恒_本月": "0.00",
                    "阳原仁恒_累计": "0.00",
                }
            )

        internal_summary = directory / "internal.csv"
        with internal_summary.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(
                file,
                fieldnames=("范围", "销方主体", "购买方名称", "内部销售金额"),
            )
            writer.writeheader()
            writer.writerows(
                [
                    {"范围": "本月", "销方主体": "科技", "购买方名称": "境内集团公司", "内部销售金额": "2888295.77"},
                    {"范围": "本月", "销方主体": "阳原", "购买方名称": "境内集团公司", "内部销售金额": "6149245.63"},
                    {"范围": "累计", "销方主体": "科技", "购买方名称": "境内集团公司", "内部销售金额": "7715537.95"},
                    {"范围": "累计", "销方主体": "科技", "购买方名称": WORKFLOW.GROUP_COMPANIES[WORKFLOW.US_COMPANY], "内部销售金额": "513540.52"},
                    {"范围": "累计", "销方主体": "阳原", "购买方名称": "境内集团公司", "内部销售金额": "15207494.64"},
                ]
            )

        ledger = {
            "本月_按主体": {
                WORKFLOW.GROUP_COMPANIES["科技"]: Decimal("5222144.38"),
                WORKFLOW.GROUP_COMPANIES["阳原"]: Decimal("1619292.03"),
            },
            "累计_按主体": {
                WORKFLOW.GROUP_COMPANIES["科技"]: Decimal("16674327.82"),
                WORKFLOW.GROUP_COMPANIES["阳原"]: Decimal("2239061.94"),
            },
        }
        m7_rows = [
            {"年度类型": "本年", "范围": "累计", "差异": "266743.74"},
            {"年度类型": "本年", "范围": "本月", "差异": "271168.52"},
        ]
        return pre_summary, internal_summary, ledger, m7_rows

    def test_subject_rows_directly_reconcile_m5_to_ledger(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            inputs = self.make_inputs(Path(temp_dir))
            rows = WORKFLOW.build_current_subject_m7_breakdown(*inputs)

        self.assertEqual(
            rows,
            [
                {"范围": "本月", "主体": "仁恒科技", "主体差异": "271168.52", "对集团差异的影响": "271168.52"},
                {"范围": "本月", "主体": "阳原仁恒", "主体差异": "0.00", "对集团差异的影响": "0.00"},
                {"范围": "累计", "主体": "仁恒科技", "主体差异": "271168.52", "对集团差异的影响": "271168.52"},
                {"范围": "累计", "主体": "阳原仁恒", "主体差异": "-4424.78", "对集团差异的影响": "-4424.78"},
                {"范围": "累计合计", "主体": "—", "主体差异": "—", "对集团差异的影响": "266743.74"},
            ],
        )

    def test_unallocated_group_difference_is_not_forced_to_a_subject(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pre, internal, ledger, m7_rows = self.make_inputs(Path(temp_dir))
            next(row for row in m7_rows if row["范围"] == "本月")["差异"] = "271268.52"
            rows = WORKFLOW.build_current_subject_m7_breakdown(pre, internal, ledger, m7_rows)

        self.assertIn(
            {"范围": "本月", "主体": "其他主体/未定位", "主体差异": "100.00", "对集团差异的影响": "100.00"},
            rows,
        )

    def test_review_workbook_contains_subject_breakdown_sheet(self):
        rows = []
        for index, (year_type, scope, difference) in enumerate(
            (("本年", "累计", "266743.74"), ("去年同期", "累计", "0.00"), ("本年", "本月", "271168.52"), ("去年同期", "本月", "0.00")),
            start=1,
        ):
            rows.append(
                {
                    "差异ID": f"M7-{index:02d}",
                    "年度类型": year_type,
                    "范围": scope,
                    "M5抵销后营业收入": difference,
                    "销售台账金额（不含税）": "0.00",
                    "差异": difference,
                    "无票销售金额": "0.00",
                    "无票后剩余差异": difference,
                    "无票归因": "无票销售未解释",
                }
            )
        subject_rows = [
            {"范围": "本月", "主体": "仁恒科技", "主体差异": "271168.52", "对集团差异的影响": "271168.52"},
            {"范围": "本月", "主体": "阳原仁恒", "主体差异": "0.00", "对集团差异的影响": "0.00"},
            {"范围": "累计", "主体": "仁恒科技", "主体差异": "271168.52", "对集团差异的影响": "271168.52"},
            {"范围": "累计", "主体": "阳原仁恒", "主体差异": "-4424.78", "对集团差异的影响": "-4424.78"},
            {"范围": "累计合计", "主体": "—", "主体差异": "—", "对集团差异的影响": "266743.74"},
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "review.xlsx"
            GRAPH_WORKFLOW.write_difference_review_workbook(
                output,
                rows,
                {},
                subject_rows,
                {
                    "year": 2026,
                    "month": 2,
                    "current_input_fingerprint": "current",
                    "prior_input_fingerprint": "prior",
                },
                1,
            )
            workbook = load_workbook(output, data_only=False)
            sheet = workbook["主体差异拆分"]
            values = list(sheet.values)
            self.assertEqual(values[0], ("范围", "主体", "主体差异", "对集团差异的影响"))
            self.assertEqual(values[1], ("本月", "仁恒科技", 271168.52, 271168.52))
            self.assertEqual(values[5], ("累计合计", "—", "—", 266743.74))
            self.assertEqual(sheet["C2"].number_format, "+#,##0.00;-#,##0.00;0.00")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
