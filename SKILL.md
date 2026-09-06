---
name: weekly-report
description: Use when the user mentions writing, updating, creating, or filling out their weekly work report, work summary, or 周报. Triggers on phrases like "写周报", "更新周报", "weekly report", "工作总结", or when the user describes work progress that should be formatted into the company's Excel weekly report template.
---

# 周报助手 (Weekly Report Assistant for Pi Agent)

## Overview

通过自然语言交互，自动将用户的工作内容结构化，并**基于公司标准 Excel 模板**安全填充生成正式周报文件。

**核心原则：在收集完所有必填板块并获得用户明确确认前，严禁生成文件。**

## When to Use

- 用户说"写周报"、"更新周报"、"帮我写本周周报"、"weekly report"
- 用户贴出一周工作碎碎念，需要整理成规范的正式周报 Excel
- 用户需要基于上周的周报做顺延更新
- **When NOT to use**: 用户只是询问 Excel 技巧、普通日程管理、或无需生成周报的日常闲聊

---

## ⚠️ CRITICAL RULES (绝不违反的铁律)

1. **严禁直接编辑 .xlsx 二进制文件**：
   - Excel 文件是复杂的二进制压缩包（OOXML）。**绝对不可使用 `edit` 或 `write` 工具直接修改 `.xlsx` 文件**，否则必然损坏格式。
   - 所有基于模板的生成必须调用内置脚本 `scripts/generate_report.py` 完成。
2. **绝不提前生成文件**：在收集完 4 大板块（上周完成、本周重点、下周计划、饱和度与其他）并经用户预览确认前，不可调用生成命令。
3. **必须经用户确认生成**：生成文件前，必须向用户展示结构化 Markdown 预览卡片，获得用户明确回复（如"确认"、"可以"、"生成吧"）。
4. **统一文件命名**：`工作周报 - {姓名} {周期起始日期}至{周期结束日期}.xlsx`，例如 `工作周报 - 张三 2026-05-25至2026-05-29.xlsx`（严禁使用单日期命名）。
5. **饱和度统一填写「饱和」**：不做百分比量化，一律写入"饱和"二字。
6. **底层强制职场化润色**：用户的口语化描述、粗糙短语（如"改了个bug"、"写了接口"）必须经职场化润色后再排入预览与表格，严禁把口水话原样填入 Excel！严格遵守【事实真实、表达专业】原则。

---

## Required Information Checklist (必填信息清单)

生成周报前，必须确认以下信息已收集完整：

### ☐ 1. 基本信息

- [ ] 周报周期（格式统一为：`YYYY/MM/DD-YYYY/MM/DD` 或 `YYYY-MM-DD至YYYY-MM-DD`）
- [ ] 姓名（默认：`张三`）
- [ ] 部门（默认：`核心业务研发部`）

### ☐ 2. 上周工作完成情况（1-5 条，建议 2-4 条）

每条工作项包含：

- [ ] 工作内容
- [ ] 预计完成时间（YYYY-MM-DD）与实际完成时间（YYYY-MM-DD）
- [ ] 完成状态（严格限定：`已完成` / `进行中` / `未开始` / `延期`）
- [ ] 成果说明
- [ ] 遇到的问题与解决方案（如无填"暂无"）

### ☐ 3. 本周重点工作（1-5 条，建议 2-4 条）

每条工作项包含：

- [ ] 工作内容
- [ ] 优先级（严格限定：`高` / `中` / `低`）
- [ ] 计划开始时间与计划完成时间（YYYY-MM-DD）
- [ ] 责任人（默认本人）
- [ ] **协作人（必须主动询问，严禁遗漏！）**
- [ ] 遇到的问题与预计完成时间

### ☐ 4. 下周工作计划（1-5 条，建议 2-4 条）

每条工作项包含：

- [ ] 工作内容
- [ ] 优先级（`高` / `中` / `低`）
- [ ] 计划开始时间与完成时间
- [ ] 责任人与**协作人（必须询问）**

### ☐ 5. 工作饱和度评估与其他说明

- [x] 整体工作饱和度：**统一填写「饱和」**
- [x] 工作负荷说明：**由 AI 根据本周多线工作内容自动提炼总结**
- [ ] 需协调支持的资源/事项（必须向用户确认）
- [ ] 其他工作说明/建议（可选）

---

## 职场专业化智能润色规范 (Core Foundation: Professional Polishing)

所有用户输入的口语化内容，Agent **必须在底层自动进行规范化职场提炼**后再呈递预览与填入 Excel。

### 1. 润色公式

$$\text{【精准动作动词】} + \text{【业务系统/功能模块】} + \text{【具体手段/工作重心】} + \text{【阶段成果/业务价值】}$$

### 2. 经典转换对照表

| 用户口语描述 | 职场专业化润色表达 |
| --- | --- |
| "改了结算流水线并发的bug" | 排查并修复订单结算流水线并发冲突异常，优化分布式锁边界校验，确保交易幂等无误 |
| "跟产品对了多租户权限二期" | 参与用户中心多租户权限系统业务需求评审，对齐核心角色鉴权规则与验收标准，完成技术任务拆解 |
| "写了几个接口" | 完成XX核心业务模块 RESTful 接口开发与联调，通过接口幂等性验证与单元测试 |
| "线上有点慢，去查了下" | 排查并定位生产环境服务响应延迟瓶颈，优化慢查询 SQL 索引与缓存策略，提升接口响应效率 |
| "写周报脚本" | 封装周报自动化生成 CLI 工具链，打通数据结构校验与模板安全填充，提升填报效率 |
| "下周准备测一测发版" | 推进系统联调与核心业务流回归测试，协同测试与业务团队完成灰度发布前验证 |

### 3. 成果说明自动推导

用户若未提供详细「成果说明」，AI 需结合工作内容自动给出合理的正规成果（如“完成代码实现并通过自测”、“达成需求一致性结论”、“消除潜在资损隐患”），并在预览中向用户确认。

### 4. 真实性底线

- 润色的是**表达精度与汇报质感**，**绝不可编造未发生的功能、虚构的数据指标或不存在的外部协作**。
- 如遇用户描述极度模糊（如"弄了一下系统"），必须先追问是哪个系统/什么动作，再行润色。

---

## 智能交互流程与数据流

### 流程 A：全新周报填报（向导或自由式）

1. **信息收集**：
   - 确认周期与基本信息。
   - 逐项询问上周、本周、下周内容（必须确认协作人！）。
2. **预览展示**：
   向用户输出结构清晰的 Markdown 预览：

   ```markdown
   ### 📋 周报预览
   - **周期**：2026/05/25 - 2026/05/29
   - **姓名/部门**：张三 / 核心业务研发部
   
   #### 1. 上周完成情况
   - [已完成] 用户中心多租户权限系统重构（成果：完成评审，无阻塞）
   
   #### 2. 本周重点工作
   - [高] 高并发订单结算流水线性能优化（协作人：李四；问题：并发锁边界调优）
   
   #### 3. 下周工作计划
   - [高] 异步消息通知网关开发（协作人：李四）
   
   #### 4. 饱和度评估与其他
   - 饱和度：饱和（双线推进，进度符合预期）
   - 需协调资源：需业务部门及时确认结算流水线接口规范
   
   👉 请确认以上内容无误，回复「确认」后我将基于 Excel 模板为您生成周报。
   ```

3. **安全生成文件（数据流协议）**：
   用户回复确认后：
   - **步骤 1**：使用 Pi 的 `write` 工具将数据存入临时 JSON 文件（如 `.weekly_report_payload.json`）。
   - **步骤 2**：使用 `bash` 工具调用脚本进行 Excel 填充：

     ```bash
     python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\generate_report.py" .weekly_report_payload.json
     ```

   - **步骤 3**：脚本返回 JSON 结果，如：
     `{"status": "success", "file": "C:\\Users\\Administrator\\Desktop\\周报\\工作周报 - 张三 2026-05-25至2026-05-29.xlsx"}`
   - **步骤 4**：清理临时 JSON 文件并向用户汇报成功。

### 流程 B：基于上周周报快速顺延更新

当用户说“基于上周周报写”、“更新本周周报”时：

1. **自动提取顺延建议**：
   执行命令：

   ```bash
   python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\generate_report.py" --migrate-last-week
   ```

2. 脚本将自动扫描输出目录下的最新周报，返回：
   - 建议上周完成情况（来自上周的“本周重点”）
   - 建议本周重点工作（来自上周的“下周计划”）
3. AI 呈递建议给用户，只需用户补充实际成果、调整本周细节和填补下周计划即可，大幅节省时间！

---

## 脚本接口规范 (CLI Reference)

脚本位置：`<skill_dir>/scripts/generate_report.py`

| 命令 | 用途 | 输出格式 |
| ------ | ------ | ---------- |
| `python generate_report.py <payload.json>` | 从 JSON 文件生成 Excel 周报 | JSON `{"status": "success", "file": "..."}` |
| `python generate_report.py --migrate-last-week` | 自动读取上一份周报并生成顺延建议 | JSON `{"status": "success", "data": {...}}` |
| `python generate_report.py --last-week` | 读取上一份周报全部原始内容 | JSON `{"status": "success", "data": {...}}` |
| `python generate_report.py --validate-template` | 验证周报模板完整性 | JSON `{"status": "success", ...}` |
| `python generate_report.py --test` | 运行内置示例生成测试 | JSON `{"status": "success", ...}` |

### Payload JSON 结构规范

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

---

## 模板与输出约定

- **模板位置**：随 Skill 自带 `templates/工作周报模板.xlsx`，亦可通过环境变量 `WKR_TEMPLATE_PATH` 指定。
- **输出目录**：默认为桌面周报目录 `~/Desktop/周报/`，亦可通过环境变量 `WKR_OUTPUT_DIR` 指定。
- **单元格格式保障**：脚本已封装 openpyxl 单元格写入、自动换行（`wrap_text=True`）、保留原有公式和单元格边框样式，绝不破坏模板原生视觉。
