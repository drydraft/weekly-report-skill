# 必填字段与取值约束

上级：[SKILL.md](../SKILL.md)。构造 payload 前必读；对应的 JSON 结构见 [cli-and-payload.md](cli-and-payload.md)。

## 1. 基本信息

| 字段 | 是否必填 | 默认值 | 约束 |
| --- | --- | --- | --- |
| `period` | 必填 | — | 周期，见下方格式表 |
| `name` | 必填 | `张三` | 写入表头姓名格 |
| `department` | 必填 | `核心业务研发部` | 写入表头部门格 |

### 周期格式

`period` 支持以下写法，脚本会统一解析成 `YYYY-MM-DD至YYYY-MM-DD` 作为文件名后缀：

| 输入写法 | 例 |
| --- | --- |
| 斜杠 + 短横 | `2026/05/25-2026/05/29` |
| 短横 + 空格分隔 | `2026-05-26 - 2026-05-30` |
| 中文「至」 | `2026-05-26至2026-05-30` |

**不要在 `period` 里放单日期**，也**不要自己拼文件名**——文件名由脚本生成。

## 2. 四大板块必填清单

### ☐ 上周工作完成情况 `last_week`（1–5 条，建议 2–4 条）

| 字段 | 必填 | 取值 |
| --- | --- | --- |
| `content` | ✅ | 工作内容（职场化表达） |
| `planned_date` | ✅ | `YYYY-MM-DD` |
| `actual_date` | ✅ | `YYYY-MM-DD` |
| `status` | ✅ | **仅限** `已完成` / `进行中` / `未开始` / `延期` |
| `results` | ✅ | 成果说明 |
| `problems` | ✅ | 遇到的问题，无则填 `暂无` |
| `solutions` | ➖ | 解决方案，可空 |
| `notes` | ➖ | 备注，可空 |

### ☐ 本周重点工作 `this_week`（1–5 条，建议 2–4 条）

| 字段 | 必填 | 取值 |
| --- | --- | --- |
| `content` | ✅ | 工作内容 |
| `priority` | ✅ | **仅限** `高` / `中` / `低` |
| `start_date` / `end_date` | ✅ | `YYYY-MM-DD` |
| `owner` | ✅ | 责任人，默认本人 |
| `collaborators` | ✅ | **协作人，必须主动询问，严禁遗漏或留空** |
| `problems` | ✅ | 遇到的问题，无则填 `暂无` |
| `eta` | ✅ | 预计完成时间 `YYYY-MM-DD` |
| `notes` | ➖ | 备注，可空 |

### ☐ 下周工作计划 `next_week`（1–5 条，建议 2–4 条）

字段同 `this_week`，但**没有 `problems` 列**（模板里下周计划不含「遇到的问题」）：

| 字段 | 必填 |
| --- | --- |
| `content` / `priority` / `start_date` / `end_date` / `owner` / `collaborators` / `eta` | ✅ |
| `notes` | ➖ |

### ☐ 工作饱和度评估与其他说明

| 字段 | 必填 | 说明 |
| --- | --- | --- |
| 饱和度 | — | **不由 payload 提供**：脚本固定写入「饱和」，不做百分比量化 |
| `workload_desc` | ✅ | 工作负荷说明，由 AI 依据本周多线内容提炼 |
| `resources_needed` | ✅ | 需协调支持的资源/事项，**必须向用户确认** |
| `other_notes` | ➖ | 其他说明/建议，可空 |

## 3. 条数与位置约束

- **每个板块最多 5 条**（脚本 `fill_section(max_items=5)`）。超过 5 条时按优先级合并或删减，不要指望多出来的行被写出。
- 脚本按**表头文字**动态定位区块，缺省兜底行号如下；模板结构不要改动：

| 区块 | 定位依据（A 列文字） | 数据起始行 |
| --- | --- | --- |
| 上周工作完成情况 | `上周工作完成情况` | 表头行 + 2（兜底 6） |
| 本周重要工作项 | `本周重要工作项` | 表头行 + 2（兜底 14） |
| 下周工作计划 | `下周工作计划` | 表头行 + 2（兜底 22） |
| 工作饱和度评估 | `工作饱和度评估` | 表头行 + 1（兜底 29） |

## 4. 字段 → Excel 单元格映射（排障用）

表头行 2：`B2` = period，`E2` = name，`H2` = department。

| 板块 | 字段 → 列号（A=1） |
| --- | --- |
| 上周完成 | `content`→2 `planned_date`→3 `actual_date`→4 `status`→5 `results`→7 `problems`→8 `solutions`→9 `notes`→10 |
| 本周重点 | `content`→2 `priority`→3 `start_date`→4 `end_date`→5 `owner`→6 `collaborators`→7 `problems`→8 `eta`→9 `notes`→10 |
| 下周计划 | `content`→2 `priority`→3 `start_date`→4 `end_date`→5 `owner`→6 `collaborators`→7 `eta`→8 `notes`→9 |
| 饱和度 | 数据行 `B`=`饱和`、`E`=`workload_desc`；下一行 `B`=`resources_needed`、`E`=`other_notes` |

> 示例：板块表头为 `上周工作完成情况` 且落在第 4 行 → 数据从第 6 行开始，`status` 写进 `E6`。
