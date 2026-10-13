---
name: weekly-report
description: "把一周的大白话工作碎碎念职场化提炼，填充进公司标准 Excel 周报模板并输出 .xlsx 文件。Use when the user mentions 写周报、更新周报、周报、weekly report、工作总结、填周报，或贴出一周工作内容要求整理成正式周报；也用于基于上周周报做顺延更新。Not for: 只问 Excel 操作技巧、普通日程管理、其他文档（Word/PPT/邮件）的写作。"
---

# 周报助手 (Weekly Report Assistant)

## Overview

把用户一周的碎片化描述，自动提炼为职场化表达，并**零损保真地写入公司标准 Excel 模板**，产出可直接提交的 `.xlsx`。

**分工原则：内容交给大模型提炼，文件落地交给确定性 Python 脚本。**

**核心原则：在收集完四大板块并获得用户明确确认之前，严禁生成任何文件。**

## When to Use

- 用户说「写周报」「更新周报」「帮我写本周周报」「weekly report」「填周报」
- 用户贴出一周的工作碎碎念，需要整理成规范的正式周报 Excel
- 用户需要基于上周周报做顺延更新（见 `workflows/B-migrate-last-week.md`）
- 上游 `weekly-plan` skill 走到「收口出周报」时交接过来

**When NOT to use**：只问 Excel 技巧 / 普通日程管理 / 需要产出 Word、PPT、邮件等其他格式 / 无需生成文件的日常闲聊。

---

## Execution Protocol（读完再动手）

1. **Sequential Lock**：下面的三道门必须按顺序过。上一道的 Postconditions 未满足前，不得开始下一道。
2. **State Echo**：进入每道门时输出一行状态，格式 `[门 N] <名称> — preconditions: PASS/FAIL`。
3. **No Silent Skips**：有意跳过某步（例如用户已提供全部字段）必须写明跳过原因，不得静默略过。
4. **No Fabrication**：任何应由用户提供或由脚本返回的值都必须真实取得——协作人、日期、成果说明都不许编造；缺少必要工具调用就停下来问。

---

## Iron Rules（硬门禁，任一违反即返工）

1. **严禁用 `edit` / `write` 工具直接改 `.xlsx`**。Excel 是 OOXML 二进制压缩包，直接改必然损坏格式与公式。一切写入必须经 `scripts/generate_report.py`。
2. **绝不提前生成文件**。四大板块（上周完成 / 本周重点 / 下周计划 / 饱和度与其他）未收集齐、且用户未明确确认前，不得调用生成命令。
3. **必须经用户显式确认**。生成前必须呈递 Markdown 预览卡片，拿到「确认 / 可以 / 生成吧」这类明确回复才算确认；用户改口或提新信息 → 回到门 1 重新收集。
4. **协作人必问**。每条本周重点、下周计划都要有 `collaborators`；没有就主动问，**严禁留空或默认填「无」**。「本周重点」还要补齐 `problems`（如无填「暂无」）与 `eta`。
5. **饱和度统一填「饱和」**，不做百分比量化。`workload_desc`（工作负荷说明）由 AI 依据本周多线内容提炼。
6. **文件命名固定**：`工作周报 - {姓名} {周期起始}至{周期结束}.xlsx`。严禁单日期命名，严禁自造其它格式。
7. **必须跑校验**：payload 落盘后先跑 `scripts/validate_payload.py`，通过才允许调 `generate_report.py`。校验不过 → 修 payload → 重跑，直到通过。

---

## Files — Load On Demand

**`<skill_dir>` = `~/.pi/agent/skills/weekly-report`**（PowerShell 中写作 `"$env:USERPROFILE\.pi\agent\skills\weekly-report"`）

| File | When to load |
| --- | --- |
| `references/fields.md` | 收集信息前 —— 四大板块必填清单、取值枚举、字段到 Excel 单元格的映射、条数上限 |
| `references/cli-and-payload.md` | 要构造 payload 或调用脚本前 —— Payload JSON 结构、CLI 命令表、模板与输出目录约定 |
| `references/polishing.md` | 做职场化润色前 —— 润色公式、转换对照表、成果推导与真实性底线 |
| `workflows/A-new-report.md` | 用户要写一份新周报（默认路径，**必读**） |
| `workflows/B-migrate-last-week.md` | 用户说「基于上周周报写」「更新周报」时 |
| `scripts/generate_report.py` | 直接执行，**不要读进上下文**；只在排障时才读源码 |
| `scripts/validate_payload.py` | 直接执行，**不要读进上下文** |

---

## Workflow（三道门，顺序不可换）

### 门 1 — 收集与润色

- 读 `references/fields.md` 对齐必填项，读 `references/polishing.md` 掌握润色标准。
- 逐个板块收集；用户已给的直接采用，缺的**一轮问完**（≤3 个问题，超出取最关键的三个，其余按 `fields.md` 的默认值）。
- 所有口语化内容在底层完成职场化提炼后再进预览，**禁止把口水话原样填入 Excel**。
- 走「基于上周顺延」分支时，先跑 `--migrate-last-week` 拿建议，再补齐差量。

**Postconditions**：四大板块字段齐全（含每条协作人）、周期与姓名部门确定。不满足 → 停在门 1。

### 门 2 — 预览与确认

- 按 `workflows/A-new-report.md` 的预览卡片模板输出 Markdown 预览。
- 用 `AskUserQuestion` 或直接请用户回复确认；**未拿到明确确认不得进入门 3**。
- 用户提出修改 → 改完重新出预览。

**Postconditions**：用户明确确认。不满足 → 停在门 2。

### 门 3 — 校验 → 生成 → 交付

1. 用 `write` 把 payload 写到工作目录的 `.weekly_report_payload.json`（结构与字段见 `references/cli-and-payload.md`）。
2. **先校验**：

   ```powershell
   python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\validate_payload.py" .weekly_report_payload.json
   ```

   返回 `{"status":"ok"}` 才继续；返回 `{"status":"invalid", "errors":[...]}` → 按 `errors` 修正 payload，重新校验。

3. **再生成**：

   ```powershell
   python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\generate_report.py" .weekly_report_payload.json
   ```

   返回 `{"status":"error", ...}` → 修 payload 重新生成，不得跳过。

4. 删除临时 payload 文件，向用户回报生成路径（用 `present_files` 打开或直接给绝对路径）。

**Postconditions**：脚本返回 `status: success` 且目标文件存在。

---

## Convergence（什么时候停）

- 门 1 最多追问 2 轮；仍缺的信息按 `fields.md` 默认值填充，并**显式告知用户用了默认值**。
- 校验—修正循环最多 3 轮；3 轮仍不过说明字段定义与 payload 结构不匹配，停下来向用户说明冲突点，不要硬凑。
- 生成成功后不再主动重排、重润色、重复导出，除非用户明确要求改。

---

## 排障

| 症状 | 处理 |
| --- | --- |
| 脚本报模板缺失 | 确认 `templates/工作周报模板.xlsx` 存在，或设 `WKR_TEMPLATE_PATH` |
| 输出目录写不进去 | 确认 `WKR_OUTPUT_DIR`（默认 `~/Desktop/周报/`）可写 |
| 单元格位置对不上 | 跑 `--validate-template` 看脚本识别到的区块行；模板结构被改过就恢复原模板 |
| 生成的表里样式丢失 | 说明绕过了脚本直接写 xlsx —— 回到门 3，必须用脚本生成 |
