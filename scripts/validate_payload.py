#!/usr/bin/env python3
"""
周报 payload 静态校验器
=======================

在调用 generate_report.py 生成 Excel 之前跑一遍，把「字段缺失 / 枚举写错 / 条数超限 /
日期不可解析 / 协作人漏填」这类问题挡在写盘之前。

与 generate_report.py 的分工：
  validate_payload.py  只做静态校验，不依赖 openpyxl，不写任何文件
  generate_report.py   真正写 Excel

用法：
  python validate_payload.py payload.json
  python validate_payload.py -                 # 从 stdin 读
  python validate_payload.py --json '{"period": ...}'

输出（stdout 单行 JSON）：
  通过  {"status": "ok", "warnings": [...], "summary": {...}}
  不通过 {"status": "invalid", "errors": [...], "warnings": [...], "summary": {...}}

退出码：通过 0，不通过 1，入参/读取错误 2。
"""

import argparse
import json
import re
import sys
from datetime import datetime

# Windows 控制台默认用 GBK，中文错误信息会乱码；固定成 UTF-8 保证调用方按 UTF-8 解码即可读
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

MAX_ITEMS = 5

DATE_PATTERN = re.compile(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}")

LAST_WEEK_STATUS = ["已完成", "进行中", "未开始", "延期"]
PRIORITY = ["高", "中", "低"]

# section -> (字段名, 是否必填)
SECTIONS = {
    "last_week": [
        ("content", True),
        ("planned_date", True),
        ("actual_date", True),
        ("status", True),
        ("results", True),
        ("problems", True),
    ],
    "this_week": [
        ("content", True),
        ("priority", True),
        ("start_date", True),
        ("end_date", True),
        ("owner", True),
        ("collaborators", True),
        ("problems", True),
        ("eta", True),
    ],
    "next_week": [
        ("content", True),
        ("priority", True),
        ("start_date", True),
        ("end_date", True),
        ("owner", True),
        ("collaborators", True),
        ("eta", True),
    ],
}

DATE_FIELDS = {
    "last_week": ["planned_date", "actual_date"],
    "this_week": ["start_date", "end_date", "eta"],
    "next_week": ["start_date", "end_date", "eta"],
}


def _is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def _parse_date(value):
    """返回 (datetime | None, 原始文本)。仅接受 YYYY-MM-DD / YYYY/MM/DD。"""
    text = str(value).strip()
    normalized = text.replace("/", "-")
    try:
        return datetime.strptime(normalized, "%Y-%m-%d"), text
    except ValueError:
        return None, text


def format_period_suffix(period):
    """
    与 generate_report.format_period_suffix 保持一致的实现（校验器不 import 主脚本，
    以便在未安装 openpyxl 的环境里也能跑）。一致性由 tests/test_validate_payload.py 保证。
    """
    matches = DATE_PATTERN.findall(str(period))
    if len(matches) < 2:
        return datetime.now().strftime("%Y-%m-%d")

    formatted = []
    for date_text in matches[:2]:
        normalized = date_text.replace("/", "-")
        try:
            formatted.append(datetime.strptime(normalized, "%Y-%m-%d").strftime("%Y-%m-%d"))
        except ValueError:
            formatted.append(normalized)

    return f"{formatted[0]}至{formatted[1]}"


def _check_period(period, errors, warnings):
    if _is_blank(period):
        errors.append("period: 缺少周期（必填），格式如 2026/05/25-2026/05/29")
        return
    matches = DATE_PATTERN.findall(str(period))
    if len(matches) < 2:
        errors.append(
            f"period: 解析不到两个日期（当前值 {period!r}）—— 生成时文件名会静默退化成当天日期，"
            "必须写成 起始至结束 两个日期"
        )
        return
    start, end = (_parse_date(m) for m in matches[:2])
    if start[0] is None:
        errors.append(f"period: 起始日期无法解析：{start[1]!r}")
    if end[0] is None:
        errors.append(f"period: 结束日期无法解析：{end[1]!r}")
    if start[0] and end[0] and end[0] < start[0]:
        errors.append(
            f"period: 结束日期早于起始日期（{start[1]} < {end[1]}）"
        )
    if len(matches) > 2:
        warnings.append(
            f"period: 出现了 {len(matches)} 个日期，只有前两个会用于文件名"
        )


def _check_section(name, items, errors, warnings):
    if items is None:
        errors.append(f"{name}: 缺失（应为数组）")
        return 0
    if not isinstance(items, list):
        errors.append(f"{name}: 必须是数组，当前是 {type(items).__name__}")
        return 0

    if len(items) == 0:
        warnings.append(f"{name}: 为空，建议填 2-4 条（模板每板块最多 {MAX_ITEMS} 条）")
    if len(items) > MAX_ITEMS:
        errors.append(
            f"{name}: 有 {len(items)} 条，超过模板上限 {MAX_ITEMS} 条；多出来的行不会写入，请合并或删减"
        )

    for idx, item in enumerate(items):
        path = f"{name}[{idx}]"
        if not isinstance(item, dict):
            errors.append(f"{path}: 必须是对象")
            continue

        for field, required in SECTIONS[name]:
            if required and _is_blank(item.get(field)):
                if field == "collaborators":
                    errors.append(
                        f"{path}: 协作人未填 —— 铁律要求必须主动询问，不能留空带队"
                    )
                else:
                    errors.append(f"{path}: 缺少必填字段 {field}")

        if name == "last_week":
            status = item.get("status")
            if not _is_blank(status) and status not in LAST_WEEK_STATUS:
                errors.append(
                    "{}: status 只能取 {}，当前 {!r}".format(
                        path, "/".join(LAST_WEEK_STATUS), status
                    )
                )
        if name in ("this_week", "next_week"):
            priority = item.get("priority")
            if not _is_blank(priority) and priority not in PRIORITY:
                errors.append(
                    "{}: priority 只能取 {}，当前 {!r}".format(
                        path, "/".join(PRIORITY), priority
                    )
                )

        for field in DATE_FIELDS[name]:
            value = item.get(field)
            if _is_blank(value):
                continue
            parsed, text = _parse_date(value)
            if parsed is None:
                errors.append(f"{path}: {field} 不是 YYYY-MM-DD 格式：{text!r}")

        if name in ("this_week", "next_week"):
            start = item.get("start_date")
            end = item.get("end_date")
            if not _is_blank(start) and not _is_blank(end):
                s, _ = _parse_date(start)
                e, _ = _parse_date(end)
                if s and e and e < s:
                    errors.append(
                        f"{path}: end_date 早于 start_date（{end} < {start}）"
                    )
            eta = item.get("eta")
            end = item.get("end_date")
            if not _is_blank(eta) and not _is_blank(end):
                eta_d, _ = _parse_date(eta)
                end_d, _ = _parse_date(end)
                if eta_d and end_d and eta_d > end_d:
                    warnings.append(
                        f"{path}: eta（{eta}）晚于 end_date（{end}），确认是否是笔误"
                    )

    return len(items)


def validate_payload(data):
    """返回 (errors, warnings, summary)。"""
    errors = []
    warnings = []

    if not isinstance(data, dict):
        return ["payload 顶层必须是 JSON 对象"], [], {}

    unknown = sorted(set(data) - {
        "period", "name", "department", "last_week", "this_week", "next_week",
        "workload_desc", "resources_needed", "other_notes", "saturation",
    })
    if unknown:
        warnings.append(f"存在未识别的顶层字段（会被忽略）：{', '.join(unknown)}")

    if "saturation" in data:
        warnings.append(
            "saturation: 该字段已废弃 —— 饱和度现在固定写入「饱和」，请删除该字段"
        )

    _check_period(data.get("period"), errors, warnings)

    for field in ("name", "department"):
        if _is_blank(data.get(field)):
            warnings.append(f"{field}: 未填，将使用模板默认值")

    for field in ("workload_desc", "resources_needed"):
        if _is_blank(data.get(field)):
            warnings.append(f"{field}: 未填（建议填写，饱和度板块会留空）")

    counts = {}
    for name in SECTIONS:
        counts[name] = _check_section(name, data.get(name), errors, warnings)

    suffix = None
    if not _is_blank(data.get("period")):
        suffix = format_period_suffix(data.get("period"))
    summary = {
        "period_suffix": suffix,
        "expected_filename": (
            "工作周报 - {} {}.xlsx".format(data.get("name") or "张三", suffix)
            if suffix else None
        ),
        "counts": counts,
    }
    return errors, warnings, summary


def load_payload(args):
    if args.json_text is not None:
        return json.loads(args.json_text)
    if args.input_file == "-":
        return json.loads(sys.stdin.read())
    with open(args.input_file, "r", encoding="utf-8") as fh:
        return json.load(fh)


def main(argv=None):
    parser = argparse.ArgumentParser(description="周报 payload 静态校验器")
    parser.add_argument("input_file", nargs="?", help="payload JSON 文件路径（'-' 从 stdin 读）")
    parser.add_argument("--json", dest="json_text", help="内联 JSON 字符串")
    args = parser.parse_args(argv)

    if not args.input_file and args.json_text is None:
        parser.print_help()
        return 2

    try:
        data = load_payload(args)
    except FileNotFoundError:
        print(json.dumps(
            {"status": "error", "message": f"payload 文件不存在: {args.input_file}"},
            ensure_ascii=False))
        return 2
    except json.JSONDecodeError as exc:
        print(json.dumps(
            {"status": "error", "message": f"JSON 解析失败: {exc}"},
            ensure_ascii=False))
        return 2

    errors, warnings, summary = validate_payload(data)
    result = {
        "status": "invalid" if errors else "ok",
        "errors": errors,
        "warnings": warnings,
        "summary": summary,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
