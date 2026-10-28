import importlib.util
import sys
import unittest
from pathlib import Path


SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
SPEC = importlib.util.spec_from_file_location(
    "langgraph_profit_workflow", SCRIPTS / "langgraph_profit_workflow.py"
)
WORKFLOW = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = WORKFLOW
SPEC.loader.exec_module(WORKFLOW)


class PeriodPolicyTests(unittest.TestCase):
    def test_july_uses_q2_cumulative_but_not_us_monthly(self):
        policy = WORKFLOW.determine_period_policy_values(2026, 7)
        self.assertTrue(policy["us_required"])
        self.assertEqual(policy["us_confirmation_month"], 6)
        self.assertFalse(policy["include_us_monthly"])
        self.assertTrue(policy["include_us_cumulative"])
        self.assertEqual(policy["xinjiang_mode"], "ZERO")

    def test_quarter_end_includes_us_monthly_and_cumulative(self):
        policy = WORKFLOW.determine_period_policy_values(2026, 6)
        self.assertEqual(policy["us_confirmation_month"], 6)
        self.assertTrue(policy["include_us_monthly"])
        self.assertTrue(policy["include_us_cumulative"])
        self.assertEqual(policy["us_quarter_start"], 4)

    def test_2025_first_five_months_ignore_us(self):
        policy = WORKFLOW.determine_period_policy_values(2025, 5)
        self.assertFalse(policy["us_required"])
        self.assertTrue(policy["ignore_us"])
        self.assertIsNone(policy["us_confirmation_month"])


class RoutingAndCorrectionTests(unittest.TestCase):
    def test_zero_difference_routes_directly_to_publish(self):
        self.assertEqual(
            WORKFLOW.route_after_m7({"has_difference": False}), "publish_report"
        )

    def test_nonzero_difference_routes_to_review(self):
        self.assertEqual(
            WORKFLOW.route_after_m7({"has_difference": True}),
            "prepare_difference_review",
        )

    def test_only_nonzero_scopes_require_corrected_revenue(self):
        state = {
            "m7_differences": [
                {"年度类型": "本年", "范围": "本月", "差异": "0.00"},
                {"年度类型": "本年", "范围": "累计", "差异": "12.30"},
            ]
        }
        correction = WORKFLOW._validate_correction(
            state,
            {
                "reviewed_by": "财务人员",
                "reason": "确认正确累计收入",
                "corrected_revenue": {"current_cumulative": "1234.56"},
            },
        )
        self.assertEqual(
            correction["corrected_revenue"]["current_cumulative"], "1234.56"
        )
        self.assertIsNone(correction["corrected_revenue"]["current_month"])


if __name__ == "__main__":
    unittest.main()
