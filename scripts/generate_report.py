#!/usr/bin/env python3
"""
周报助手脚本 - 读取模板和生成周报 (Pi Agent 适配版)
"""

import argparse
import glob
import json
import os
import re
import sys
import warnings
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment

# 屏蔽 openpyxl 读取复合格式与条件格式时的非关键告警
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)

DEFAULT_TEMPLATE_PATH = os.path.join(SKILL_DIR, "templates", "工作周报模板.xlsx")
DEFAULT_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "周报")


def get_template_path():
    """多级查找模板路径：环境变量 -> Skill 自带模板 -> 用户常见桌面路径"""
    env_path = os.environ.get("WKR_TEMPLATE_PATH")
    if env_path and os.path.exists(env_path):
        return os.path.abspath(env_path)
    if os.path.exists(DEFAULT_TEMPLATE_PATH):
        return os.path.abspath(DEFAULT_TEMPLATE_PATH)
    desktop_template = os.path.join(
        os.path.expanduser("~"), "Desktop", "周报", "模板", "工作周报模板.xlsx"
    )
    if os.path.exists(desktop_template):
        return os.path.abspath(desktop_template)
    return os.path.abspath(DEFAULT_TEMPLATE_PATH)


def get_output_dir():
    """获取周报输出目录"""
    output_dir = os.environ.get("WKR_OUTPUT_DIR", DEFAULT_OUTPUT_DIR)
    return os.path.abspath(os.path.expanduser(output_dir))


def format_period_suffix(period):
    """从周期文本中提取起止日期并格式化为文件名后缀，如 2026-05-25至2026-05-29"""
    matches = re.findall(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", str(period))
    if len(matches) < 2:
        return datetime.now().strftime("%Y-%m-%d")

    formatted = []
    for date_text in matches[:2]:
        normalized = date_text.replace("/", "-")
        try:
            date_value = datetime.strptime(normalized, "%Y-%m-%d")
            formatted.append(date_value.strftime("%Y-%m-%d"))
        except ValueError:
            formatted.append(normalized)

    return f"{formatted[0]}至{formatted[1]}"


def set_cell_value(sheet, row, col, value):
    """设置单元格值，保留换行与上对齐格式"""
    cell = sheet.cell(row=row, column=col, value=value)
    if isinstance(value, str) and "\n" in value:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    return cell


def find_section_rows(sheet):
    """动态查找各板块标题行和数据起始行"""
    sections = {}
    for row in range(1, sheet.max_row + 1):
        cell_val = sheet.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str):
            if "上周工作完成情况" in cell_val:
                sections["last_week_header"] = row
                sections["last_week_data_start"] = row + 2
            elif "本周重要工作项" in cell_val:
                sections["this_week_header"] = row
                sections["this_week_data_start"] = row + 2
            elif "下周工作计划" in cell_val:
                sections["next_week_header"] = row
                sections["next_week_data_start"] = row + 2
            elif "工作饱和度评估" in cell_val:
                sections["saturation_header"] = row
                sections["saturation_data"] = row + 1

    # 兜底默认行号
    sections.setdefault("last_week_data_start", 6)
    sections.setdefault("this_week_data_start", 14)
    sections.setdefault("next_week_data_start", 22)
    sections.setdefault("saturation_data", 29)
    return sections


def fill_section(sheet, start_row, items, columns, max_items=5):
    """
    填充板块数据并更新序号
    """
    for i in range(max_items):
        row = start_row + i
        if i < len(items):
            item = items[i]
            # 确保序号存在
            sheet.cell(row=row, column=1, value=i + 1)
            for field, col in columns.items():
                if field in item and item[field] is not None:
                    val = item[field]
                    # 日期格式兼容
                    if isinstance(val, (datetime,)):
                        val = val.strftime("%Y-%m-%d")
                    set_cell_value(sheet, row, col, val)
        else:
            # 清空超出的空行（如果模板中有残留）
            pass


def read_last_week_report():
    """读取输出目录下最新的周报文件"""
    output_dir = get_output_dir()
    if not os.path.exists(output_dir):
        return None
    pattern = os.path.join(output_dir, "*周报 - *.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def parse_report_file(file_path):
    """解析已生成的周报文件，提取完整结构化数据"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件未找到: {file_path}")

    wb = load_workbook(file_path, data_only=True)
    if "员工工作周报" not in wb.sheetnames:
        raise ValueError(f"Excel 中未找到 '员工工作周报' 工作表: {file_path}")

    sheet = wb["员工工作周报"]
    sections = find_section_rows(sheet)

    period = sheet["B2"].value or ""
    name = sheet["E2"].value or "张三"
    department = sheet["H2"].value or "核心业务研发部"

    def extract_rows(start_row, col_map):
        rows = []
        for r in range(start_row, start_row + 5):
            content = sheet.cell(row=r, column=2).value
            if content and str(content).strip():
                item = {}
                for k, c in col_map.items():
                    val = sheet.cell(row=r, column=c).value
                    if isinstance(val, datetime):
                        val = val.strftime("%Y-%m-%d")
                    item[k] = val if val is not None else ""
                rows.append(item)
        return rows

    last_week_cols = {
        "content": 2,
        "planned_date": 3,
        "actual_date": 4,
        "status": 5,
        "results": 7,
        "problems": 8,
        "solutions": 9,
        "notes": 10,
    }
    this_week_cols = {
        "content": 2,
        "priority": 3,
        "start_date": 4,
        "end_date": 5,
        "owner": 6,
        "collaborators": 7,
        "problems": 8,
        "eta": 9,
        "notes": 10,
    }
    next_week_cols = {
        "content": 2,
        "priority": 3,
        "start_date": 4,
        "end_date": 5,
        "owner": 6,
        "collaborators": 7,
        "eta": 8,
        "notes": 9,
    }

    sat_row = sections["saturation_data"]
    return {
        "file": os.path.abspath(file_path),
        "period": str(period),
        "name": str(name),
        "department": str(department),
        "last_week": extract_rows(sections["last_week_data_start"], last_week_cols),
        "this_week": extract_rows(sections["this_week_data_start"], this_week_cols),
        "next_week": extract_rows(sections["next_week_data_start"], next_week_cols),
        "saturation": sheet.cell(row=sat_row, column=2).value or "饱和",
        "workload_desc": sheet.cell(row=sat_row, column=5).value or "",
        "resources_needed": sheet.cell(row=sat_row + 1, column=2).value or "",
        "other_notes": sheet.cell(row=sat_row + 1, column=5).value or "",
    }


def migrate_from_last_week(last_week_path):
    """
    从上周周报迁移数据：
    - 上周的「本周重要工作项」 -> 本次周报的「上周工作完成情况」（状态默认'已完成'）
    - 上周的「下周工作计划」 -> 本次周报的「本周重要工作项」
    """
    prev = parse_report_file(last_week_path)
    migrated_last_week = []
    for item in prev.get("this_week", []):
        migrated_last_week.append(
            {
                "content": item.get("content", ""),
                "planned_date": item.get("end_date") or item.get("start_date") or "",
                "actual_date": item.get("eta") or item.get("end_date") or "",
                "status": "已完成",
                "results": "",
                "problems": item.get("problems", ""),
                "solutions": "",
                "notes": item.get("notes", ""),
            }
        )

    migrated_this_week = []
    for item in prev.get("next_week", []):
        migrated_this_week.append(
            {
                "content": item.get("content", ""),
                "priority": item.get("priority", "中"),
                "start_date": item.get("start_date", ""),
                "end_date": item.get("end_date", ""),
                "owner": item.get("owner", prev.get("name", "张三")),
                "collaborators": item.get("collaborators", ""),
                "problems": "",
                "eta": item.get("eta", ""),
                "notes": item.get("notes", ""),
            }
        )

    return {
        "source_file": os.path.abspath(last_week_path),
        "previous_period": prev.get("period", ""),
        "name": prev.get("name", "张三"),
        "department": prev.get("department", "核心业务研发部"),
        "suggested_last_week": migrated_last_week,
        "suggested_this_week": migrated_this_week,
    }


def validate_template(template_path=None):
    """验证模板文件是否可用"""
    p = template_path or get_template_path()
    if not os.path.exists(p):
        return False, f"模板文件不存在: {p}"
    try:
        wb = load_workbook(p)
        if "员工工作周报" not in wb.sheetnames:
            return False, f"模板中缺少 '员工工作周报' 工作表: {p}"
        sections = find_section_rows(wb["员工工作周报"])
        return True, {
            "template_path": os.path.abspath(p),
            "sheets": wb.sheetnames,
            "sections": sections,
        }
    except Exception as e:
        return False, f"读取模板异常: {str(e)}"


def generate_weekly_report(report_data):
    """
    根据给定的 report_data 填充 Excel 模板并输出文件。
    """
    template_path = get_template_path()
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"周报模板未找到，请检查路径: {template_path}")

    wb = load_workbook(template_path)
    if "员工工作周报" not in wb.sheetnames:
        raise ValueError("周报模板缺少工作表 '员工工作周报'")

    sheet = wb["员工工作周报"]
    sections = find_section_rows(sheet)

    # 1. 表头 (Row 2)
    sheet["B2"] = report_data.get("period", "")
    sheet["E2"] = report_data.get("name", "张三")
    sheet["H2"] = report_data.get("department", "核心业务研发部")

    # 2. 上周完成情况
    last_week_cols = {
        "content": 2,
        "planned_date": 3,
        "actual_date": 4,
        "status": 5,
        "results": 7,
        "problems": 8,
        "solutions": 9,
        "notes": 10,
    }
    fill_section(
        sheet,
        sections["last_week_data_start"],
        report_data.get("last_week", []),
        last_week_cols,
    )

    # 3. 本周重点工作
    this_week_cols = {
        "content": 2,
        "priority": 3,
        "start_date": 4,
        "end_date": 5,
        "owner": 6,
        "collaborators": 7,
        "problems": 8,
        "eta": 9,
        "notes": 10,
    }
    fill_section(
        sheet,
        sections["this_week_data_start"],
        report_data.get("this_week", []),
        this_week_cols,
    )

    # 4. 下周计划 (下周无"遇到的问题"列)
    next_week_cols = {
        "content": 2,
        "priority": 3,
        "start_date": 4,
        "end_date": 5,
        "owner": 6,
        "collaborators": 7,
        "eta": 8,
        "notes": 9,
    }
    fill_section(
        sheet,
        sections["next_week_data_start"],
        report_data.get("next_week", []),
        next_week_cols,
    )

    # 5. 工作饱和度评估与其他说明
    sat_row = sections["saturation_data"]
    # 规则：饱和度统一填写「饱和」，不做百分比量化
    sheet.cell(row=sat_row, column=2, value="饱和")
    sheet.cell(row=sat_row, column=5, value=report_data.get("workload_desc", ""))
    sheet.cell(row=sat_row + 1, column=2, value=report_data.get("resources_needed", ""))
    sheet.cell(row=sat_row + 1, column=5, value=report_data.get("other_notes", ""))

    # 计算输出文件名与路径
    period = report_data.get("period", "")
    date_suffix = format_period_suffix(period)
    name = report_data.get("name", "张三")
    filename = f"工作周报 - {name} {date_suffix}.xlsx"

    output_dir = get_output_dir()
    try:
        os.makedirs(output_dir, exist_ok=True)
    except OSError as e:
        raise OSError(f"创建周报输出目录失败: {output_dir}, 错误: {e}") from e
    output_path = os.path.join(output_dir, filename)

    wb.save(output_path)
    return os.path.abspath(output_path)


def main():
    parser = argparse.ArgumentParser(description="周报助手 CLI (Pi Agent 专用)")
    parser.add_argument(
        "input_file",
        nargs="?",
        help="包含周报数据的 JSON 文件路径（传 '-' 从 stdin 读取）",
    )
    parser.add_argument("--json", dest="inline_json", help="内联 JSON 字符串")
    parser.add_argument(
        "--validate-template", action="store_true", help="校验模板完整性并输出元信息"
    )
    parser.add_argument(
        "--last-week", action="store_true", help="读取最新周报内容，以 JSON 格式输出"
    )
    parser.add_argument(
        "--migrate-last-week",
        action="store_true",
        help="从最新周报中提取顺延工作建议，以 JSON 格式输出",
    )
    parser.add_argument("--test", action="store_true", help="运行示例生成测试")

    args = parser.parse_args()

    # 1. 验证模板
    if args.validate_template:
        ok, res = validate_template()
        if ok:
            print(
                json.dumps(
                    {"status": "success", "data": res}, ensure_ascii=False, indent=2
                )
            )
            sys.exit(0)
        else:
            print(json.dumps({"status": "error", "message": res}, ensure_ascii=False))
            sys.exit(1)

    # 2. 查看上一周周报数据
    if args.last_week:
        last_file = read_last_week_report()
        if not last_file:
            print(
                json.dumps(
                    {"status": "error", "message": "未在输出目录找到历史周报文件"},
                    ensure_ascii=False,
                )
            )
            sys.exit(1)
        try:
            data = parse_report_file(last_file)
            print(
                json.dumps(
                    {"status": "success", "data": data}, ensure_ascii=False, indent=2
                )
            )
            sys.exit(0)
        except Exception as e:
            print(
                json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)
            )
            sys.exit(1)

    # 3. 从上一周周报迁移数据
    if args.migrate_last_week:
        last_file = read_last_week_report()
        if not last_file:
            print(
                json.dumps(
                    {
                        "status": "error",
                        "message": "未在输出目录找到历史周报文件，无法迁移",
                    },
                    ensure_ascii=False,
                )
            )
            sys.exit(1)
        try:
            migrated = migrate_from_last_week(last_file)
            print(
                json.dumps(
                    {"status": "success", "data": migrated},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            sys.exit(0)
        except Exception as e:
            print(
                json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)
            )
            sys.exit(1)

    # 4. 测试模式
    if args.test:
        test_data = {
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
                    "content": "高并发订单结算流水线性能优化",
                    "priority": "高",
                    "start_date": "2026-05-18",
                    "end_date": "2026-05-21",
                    "owner": "张三",
                    "collaborators": "李四",
                    "problems": "高并发分布式锁边界调优",
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
            "workload_desc": "双线并行推进，开发与需求分析进度符合预期",
            "resources_needed": "需业务部门及时确认积分接口规范",
            "other_notes": "",
        }
        try:
            out_file = generate_weekly_report(test_data)
            print(
                json.dumps(
                    {
                        "status": "success",
                        "file": out_file,
                        "message": f"测试周报生成成功: {out_file}",
                    },
                    ensure_ascii=False,
                    indent=2,
                )
            )
            sys.exit(0)
        except Exception as e:
            print(
                json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)
            )
            sys.exit(1)

    # 5. 生成报告（读取传入的 JSON）
    raw_json = None
    if args.inline_json:
        raw_json = args.inline_json
    elif args.input_file == "-":
        raw_json = sys.stdin.read()
    elif args.input_file:
        if not os.path.exists(args.input_file):
            print(
                json.dumps(
                    {
                        "status": "error",
                        "message": f"输入文件不存在: {args.input_file}",
                    },
                    ensure_ascii=False,
                )
            )
            sys.exit(1)
        try:
            with open(args.input_file, encoding="utf-8") as f:
                raw_json = f.read()
        except OSError as e:
            print(
                json.dumps(
                    {
                        "status": "error",
                        "message": f"读取输入文件失败: {args.input_file}, 错误: {e}",
                    },
                    ensure_ascii=False,
                )
            )
            sys.exit(1)

    if not raw_json:
        parser.print_help()
        sys.exit(1)

    try:
        data = json.loads(raw_json)
        out_file = generate_weekly_report(data)
        print(
            json.dumps(
                {
                    "status": "success",
                    "file": out_file,
                    "message": f"周报已成功生成: {out_file}",
                },
                ensure_ascii=False,
            )
        )
        sys.exit(0)
    except Exception as e:
        print(
            json.dumps(
                {"status": "error", "message": f"周报生成失败: {str(e)}"},
                ensure_ascii=False,
            )
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
