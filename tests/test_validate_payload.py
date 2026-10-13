import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from scripts.generate_report import format_period_suffix as report_period_suffix
from scripts.validate_payload import format_period_suffix as validate_period_suffix
from scripts.validate_payload import validate_payload

REPO_ROOT = Path(__file__).resolve().parent.parent
VALIDATOR = REPO_ROOT / "scripts" / "validate_payload.py"


def good_payload():
    return {
        "period": "2026/05/25-2026/05/29",
        "name": "张三",
        "department": "核心业务研发部",
        "last_week": [
            {
                "content": "用户中心多租户权限系统重构",
                "planned_date": "2026-05-26",
                "actual_date": "2026-05-26",
                "status": "已完成",
                "results": "完成需求文档撰写和评审",
                "problems": "暂无",
            }
        ],
        "this_week": [
            {
                "content": "订单结算流水线性能优化",
                "priority": "高",
                "start_date": "2026-05-18",
                "end_date": "2026-05-21",
                "owner": "张三",
                "collaborators": "李四",
                "problems": "暂无",
                "eta": "2026-05-21",
            }
        ],
        "next_week": [
            {
                "content": "异步消息通知网关开发",
                "priority": "高",
                "start_date": "2026-05-20",
                "end_date": "2026-06-15",
                "owner": "张三",
                "collaborators": "李四",
                "eta": "2026-06-15",
            }
        ],
        "workload_desc": "本周双线并行推进",
        "resources_needed": "需业务部门确认接口规范",
    }


class ValidatePayloadTest(unittest.TestCase):
    def assert_ok(self, payload):
        errors, warnings, summary = validate_payload(payload)
        self.assertEqual(errors, [], f"unexpected errors: {errors}")
        return warnings, summary

    def assert_error_contains(self, payload, fragment):
        errors, _warnings, _summary = validate_payload(payload)
        self.assertTrue(errors, f"expected an error containing {fragment!r}")
        self.assertTrue(
            any(fragment in e for e in errors),
            f"no error contains {fragment!r}; got {errors}",
        )

    def test_valid_payload_passes(self):
        warnings, summary = self.assert_ok(good_payload())
        self.assertEqual(warnings, [])
        self.assertEqual(
            summary["expected_filename"], "工作周报 - 张三 2026-05-25至2026-05-29.xlsx"
        )

    def test_missing_collaborators_is_an_error(self):
        payload = good_payload()
        payload["this_week"][0]["collaborators"] = ""
        self.assert_error_contains(payload, "协作人未填")

    def test_missing_required_field_is_an_error(self):
        payload = good_payload()
        del payload["last_week"][0]["results"]
        self.assert_error_contains(payload, "缺少必填字段 results")

    def test_bad_status_enum_is_an_error(self):
        payload = good_payload()
        payload["last_week"][0]["status"] = "搞定"
        self.assert_error_contains(payload, "status 只能取")

    def test_bad_priority_enum_is_an_error(self):
        payload = good_payload()
        payload["next_week"][0]["priority"] = "紧急"
        self.assert_error_contains(payload, "priority 只能取")

    def test_bad_date_format_is_an_error(self):
        payload = good_payload()
        payload["this_week"][0]["eta"] = "下周三"
        self.assert_error_contains(payload, "不是 YYYY-MM-DD 格式")

    def test_period_with_single_date_is_an_error(self):
        payload = good_payload()
        payload["period"] = "2026-05-25"
        self.assert_error_contains(payload, "解析不到两个日期")

    def test_period_with_reversed_range_is_an_error(self):
        payload = good_payload()
        payload["period"] = "2026-05-29至2026-05-25"
        self.assert_error_contains(payload, "结束日期早于起始日期")

    def test_end_date_before_start_date_is_an_error(self):
        payload = good_payload()
        payload["this_week"][0]["end_date"] = "2026-05-01"
        self.assert_error_contains(payload, "end_date 早于 start_date")

    def test_more_than_five_items_is_an_error(self):
        payload = good_payload()
        payload["this_week"] = payload["this_week"] * 6
        self.assert_error_contains(payload, "超过模板上限")

    def test_deprecated_saturation_field_warns(self):
        payload = good_payload()
        payload["saturation"] = 0.9
        warnings, _summary = self.assert_ok(payload)
        self.assertTrue(any("saturation" in w for w in warnings), warnings)

    def test_empty_section_warns_but_passes(self):
        payload = good_payload()
        payload["next_week"] = []
        warnings, _summary = self.assert_ok(payload)
        self.assertTrue(any("next_week: 为空" in w for w in warnings), warnings)

    def test_period_suffix_matches_generate_report(self):
        for period in [
            "2026/05/25-2026/05/29",
            "2026-05-26 - 2026-05-30",
            "2026-05-26至2026-05-30",
            "2026/5/26-2026/5/30",
        ]:
            self.assertEqual(
                validate_period_suffix(period),
                report_period_suffix(period),
                f"period suffix mismatch for {period!r}",
            )


class ValidatorCliTest(unittest.TestCase):
    def _run(self, payload):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "payload.json"
            path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            return subprocess.run(
                [sys.executable, str(VALIDATOR), str(path)],
                capture_output=True,
                text=True,
                encoding="utf-8",
            )

    def test_cli_exit_zero_on_valid_payload(self):
        result = self._run(good_payload())
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(json.loads(result.stdout)["status"], "ok")

    def test_cli_exit_one_on_invalid_payload(self):
        payload = good_payload()
        payload["this_week"][0]["collaborators"] = ""
        result = self._run(payload)
        self.assertEqual(result.returncode, 1)
        body = json.loads(result.stdout)
        self.assertEqual(body["status"], "invalid")
        self.assertTrue(body["errors"])

    def test_cli_reports_missing_file(self):
        result = subprocess.run(
            [sys.executable, str(VALIDATOR), "does-not-exist.json"],
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        self.assertEqual(result.returncode, 2)
        self.assertEqual(json.loads(result.stdout)["status"], "error")


if __name__ == "__main__":
    unittest.main()
