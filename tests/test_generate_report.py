import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from scripts.generate_report import (
    generate_weekly_report,
    migrate_from_last_week,
    parse_report_file,
    read_last_week_report,
)


class GenerateWeeklyReportTest(unittest.TestCase):
    def setUp(self):
        self._old_output_dir = os.environ.get("WKR_OUTPUT_DIR")
        self.output_dir = tempfile.TemporaryDirectory()
        os.environ["WKR_OUTPUT_DIR"] = self.output_dir.name

    def tearDown(self):
        self.output_dir.cleanup()
        if self._old_output_dir is None:
            os.environ.pop("WKR_OUTPUT_DIR", None)
        else:
            os.environ["WKR_OUTPUT_DIR"] = self._old_output_dir

    def test_saturation_is_always_written_as_saturated(self):
        path = generate_weekly_report(
            {
                "period": "2026/5/26-2026/5/30",
                "name": "Tester",
                "department": "IT信息技术部",
                "last_week": [],
                "this_week": [],
                "next_week": [],
                "saturation": 0.9,
            }
        )

        sheet = load_workbook(path)["员工工作周报"]
        self.assertEqual(sheet["B29"].value, "饱和")

    def test_output_directory_is_created_when_missing(self):
        missing_dir = Path(self.output_dir.name) / "nested" / "reports"
        os.environ["WKR_OUTPUT_DIR"] = str(missing_dir)

        path = generate_weekly_report(
            {
                "period": "2026/5/26-2026/5/30",
                "last_week": [],
                "this_week": [],
                "next_week": [],
            }
        )

        self.assertTrue(Path(path).exists())

    def test_read_last_week_report_uses_configured_output_dir(self):
        path = generate_weekly_report(
            {
                "period": "2026/5/26-2026/5/30",
                "last_week": [],
                "this_week": [],
                "next_week": [],
            }
        )

        self.assertEqual(read_last_week_report(), path)

    def test_iso_date_period_uses_full_range_in_filename(self):
        path = generate_weekly_report(
            {
                "period": "2026-05-26 - 2026-05-30",
                "name": "Tester",
                "last_week": [],
                "this_week": [],
                "next_week": [],
            }
        )

        self.assertIn("2026-05-26至2026-05-30", Path(path).name)

    def test_parse_and_migrate_from_last_week(self):
        report_data = {
            "period": "2026/5/20-2026/5/24",
            "name": "Tester",
            "department": "IT部",
            "last_week": [{"content": "历史任务1", "status": "已完成"}],
            "this_week": [
                {
                    "content": "本周任务1",
                    "owner": "Tester",
                    "collaborators": "Bob",
                    "priority": "高",
                }
            ],
            "next_week": [
                {
                    "content": "下周任务1",
                    "owner": "Tester",
                    "collaborators": "Alice",
                    "priority": "中",
                }
            ],
            "workload_desc": "饱和正常",
        }
        path = generate_weekly_report(report_data)
        parsed = parse_report_file(path)
        self.assertEqual(parsed["name"], "Tester")
        self.assertEqual(len(parsed["this_week"]), 1)
        self.assertEqual(parsed["this_week"][0]["content"], "本周任务1")

        migrated = migrate_from_last_week(path)
        self.assertEqual(len(migrated["suggested_last_week"]), 1)
        self.assertEqual(migrated["suggested_last_week"][0]["content"], "本周任务1")
        self.assertEqual(len(migrated["suggested_this_week"]), 1)
        self.assertEqual(migrated["suggested_this_week"][0]["content"], "下周任务1")

    def test_cli_execution_with_json_file(self):
        test_payload = {
            "period": "2026/6/1-2026/6/5",
            "name": "CLI Tester",
            "department": "IT部",
            "last_week": [],
            "this_week": [],
            "next_week": [],
        }
        json_file = Path(self.output_dir.name) / "payload.json"
        json_file.write_text(json.dumps(test_payload), encoding="utf-8")

        script_path = (
            Path(__file__).resolve().parent.parent / "scripts" / "generate_report.py"
        )
        env = dict(os.environ)
        result = subprocess.run(
            [sys.executable, str(script_path), str(json_file)],
            capture_output=True,
            text=True,
            encoding="utf-8",  # 脚本固定以 UTF-8 输出 JSON，不跟随 Windows 控制台代码页
            env=env,
        )
        self.assertEqual(result.returncode, 0)
        res_data = json.loads(result.stdout)
        self.assertEqual(res_data.get("status"), "success")
        self.assertTrue(Path(res_data.get("file")).exists())


if __name__ == "__main__":
    unittest.main()
