# CLI 与 Payload 规范

上级：[SKILL.md](../SKILL.md)。脚本位置：`<skill_dir>/scripts/generate_report.py`；
**`<skill_dir>` = `~/.pi/agent/skills/weekly-report`**（PowerShell 写作 `"$env:USERPROFILE\.pi\agent\skills\weekly-report"`）。

脚本直接执行，**不要读进上下文**；只有排障时才读源码。

## 1. 命令表

| 命令 | 用途 | 输出 |
| --- | --- | --- |
| `python generate_report.py <payload.json>` | 从 JSON 文件生成 Excel | `{"status":"success","file":"..."}` |
| `python generate_report.py -` | 从 stdin 读 payload | 同上 |
| `python generate_report.py --json '<json>'` | 内联 JSON 直接生成 | 同上 |
| `python generate_report.py --validate-template` | 校验模板完整性并回显识别到的区块行 | `{"status":"success","data":{...}}` |
| `python generate_report.py --last-week` | 读取输出目录里最新一份周报的全部原始内容 | `{"status":"success","data":{...}}` |
| `python generate_report.py --migrate-last-week` | 生成顺延建议（上周「本周重点」→ 上周完成；「下周计划」→ 本周重点） | `{"status":"success","data":{...}}` |
| `python generate_report.py --test` | 跑内置示例生成自检 | `{"status":"success",...}` |

失败一律返回 `{"status":"error","message":"..."}` 并以非零码退出——**看到 error 就修，不要当成成功继续**。

> **输出编码**：两个脚本都固定以 **UTF-8** 输出 JSON（不跟随 Windows 控制台代码页），
> 因此调用方/测试请显式按 UTF-8 解码 stdout（`encoding="utf-8"`），否则中文路径与消息会解码失败。

## 2. Payload JSON 结构

字段含义与取值约束见 [fields.md](fields.md)。

```json
{
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
      "solutions": "",
      "notes": ""
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
      "problems": "一票否决功能需求变更",
      "eta": "2026-05-21",
      "notes": ""
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
      "notes": ""
    }
  ],
  "workload_desc": "本周双线并行推进，开发与需求分析进度符合预期",
  "resources_needed": "需业务部门及时确认结算流水线接口规范",
  "other_notes": ""
}
```

要点：

- **不存在 `saturation` 字段**。饱和度写死在模板里为「饱和」；历史版本传过 `saturation: 0.9`，现在会被忽略。
- `collaborators` 为空字符串会被校验器判为不合格（见 `scripts/validate_payload.py`）。
- 日期既接受 `YYYY-MM-DD` 字符串，也接受 Excel 的 datetime 对象（脚本会格式化）。

## 3. 输出与命名

- **文件名**：`工作周报 - {name} {start}至{end}.xlsx`，日期段由 `period` 解析得到。
  例：`工作周报 - 张三 2026-05-25至2026-05-29.xlsx`。**严禁单日期命名，严禁自己拼名字。**
- **输出目录**：默认 `~/Desktop/周报/`，不存在会自动创建。
- **写入方式**：`openpyxl` 打开模板 → 单元格级写入 → 另存新文件。模板自带的表头配色、合并单元格、公式与边框全部保留。
  含换行符的文本会自动设 `wrap_text=True` + 顶端对齐。

## 4. 环境变量

| 环境变量 | 说明 | 默认值 |
| --- | --- | --- |
| `WKR_TEMPLATE_PATH` | 自定义 Excel 模板路径 | `<skill_dir>/templates/工作周报模板.xlsx` |
| `WKR_OUTPUT_DIR` | 周报输出目录 | `~/Desktop/周报/` |

> 测试用例通过 `WKR_OUTPUT_DIR` 指向临时目录来隔离副作用，改动脚本时不要破坏这一点。

## 5. 作为模块调用

测试与外部集成直接 import 脚本里的函数（`tests/` 就是这么做的）：

```python
from scripts.generate_report import (
    generate_weekly_report,   # (report_data: dict) -> abs_path
    migrate_from_last_week,   # (last_week_path) -> dict
    parse_report_file,        # (file_path) -> dict
    read_last_week_report,    # () -> path | None
    validate_template,        # (template_path=None) -> (ok: bool, info)
)
```

因此**不要改动这些函数签名**，否则 `tests/` 会失败。
