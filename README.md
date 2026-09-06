<p align="center">
  <img src="./assets/readme/hero.svg" width="100%" alt="周报助手 Weekly Report Assistant：对 AI 讲几句大白话碎碎念，自动完成职场化提炼，并填充至公司标准 Excel 周报模板。">
</p>

<p align="center">
  <a href="https://agentskills.io"><img src="https://img.shields.io/badge/standard-Agent_Skills-0284C7?style=flat-square" alt="Agent Skills Standard"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-10B981?style=flat-square" alt="MIT License"></a>
  <a href="#-支持平台"><img src="https://img.shields.io/badge/runtime-Pi_Agent_|_OpenCode_|_Claude_Code-8B5CF6?style=flat-square" alt="Runtimes"></a>
  <a href="#-为什么需要周报助手"><img src="https://img.shields.io/badge/Excel-100%25_Fidelity-107C41?style=flat-square" alt="Excel Fidelity"></a>
</p>

---

## 💡 为什么需要周报助手？

每到周五下午，写周报往往面临两大痛点：

1. **话术包装难**：做了一堆事情，但写出来全是“改了bug”、“对了接口”，口语化太重，难以体现工作成果与业务价值；
2. **填表排版烦**：打开公司动辄几十列、带合并单元格和公式的 Excel 模板，复制粘贴经常搞坏样式、漏填协作人。

**周报助手（Weekly Report Skill）** 彻底改变这个流程 —— **把内容交给大模型提炼，把文件落地交给确定性 Python 引擎**。你只需在终端或对话框里说几句本周做的大白话，AI 自动完成职场规范化重构，并零损保真地填充进公司标准 `.xlsx` 模板，直达桌面！

<p align="center">
  <img src="./assets/readme/workflow.svg" width="100%" alt="周报助手工作流：自然语言碎碎念 → 智能职场化提炼 → 预览卡片与确认锁 → 原生 Excel 安全写入。">
</p>

---

## ✨ 核心特性

### 1. 智能职场化润色（底层核心能力）

绝不只是机械的“关键词替换”，而是依靠大模型深度理解业务上下文后，按照 **【精准动词】+【业务系统/模块】+【关键手段/重心】+【阶段价值】** 展开结构化重塑：

| 原始口语输入 | 职场规范化润色呈现（进 Excel 效果） |
| --- | --- |
| “改了结算流水线的并发bug” | **排查并修复订单结算流水线并发冲突异常，优化分布式锁边界校验，确保交易幂等无误** |
| “跟产品对了多租户权限二期” | **参与用户中心多租户权限系统业务需求评审，对齐核心角色鉴权规则与验收标准，完成技术任务拆解** |
| “写了几个接口并自测” | **完成核心业务模块 RESTful 接口开发与联调，通过接口幂等性验证与单元测试覆盖** |
| “线上有点卡，查了下” | **排查并定位生产环境服务响应延迟瓶颈，优化慢查询 SQL 索引与缓存策略，提升接口响应效率** |

### 2. 100% 原生 Excel 零损保真

* 彻底杜绝“生成一张不可选中文本的大长图”或“简单导出一个丢了样式的空 csv”。
* 内部使用 `openpyxl` 精确操作原生 `.xlsx` 单元格，**完整保留企业模板的表头配色、合并单元格、进度计算公式及换行边框**。

### 3. 跨周智能顺延继承 (`--migrate-last-week`)

* 用户只需说：“*基于上周周报写*”或“*更新周报*”。
* 脚本自动逆向扫描解析上一份周报，智能将上周的“本周重点”顺延为“上周完成”（默认已完成），将“下周计划”顺延为“本周重点”，用户只需补充少量进展即可完稿！

### 4. 严格的防漏防手抖安全锁

* **绝对安全**：必须先在对话中呈递完整的 Markdown 预览卡片，在用户明确回复“确认”、“可以”前，**绝不提前生成文件**。
* **协作人必查**：严格校验必填项，主动追问每条工作的责任人与协作人，杜绝遗漏。
* **饱和度规范**：饱和度统一标记为「饱和」，负荷说明由 AI 结合多线推进情况自动总结，体面得体。

---

## 🚀 快速开始

### 1. 安装 Skill

遵循 [Agent Skills 标准规范](https://agentskills.io)，可一键装入你最顺手的 Coding Harness：

#### 方案 A：Pi Agent (强烈推荐)

```bash
# 克隆到 Pi 全局 Skill 目录
git clone https://github.com/drydraft/weekly-report-skill.git \
  ~/.pi/agent/skills/weekly-report
```

> 在 Pi 中输入 `/reload` 即可立即就绪。

#### 方案 B：OpenCode

```bash
git clone https://github.com/drydraft/weekly-report-skill.git \
  ~/.config/opencode/superpowers/skills/weekly-report
```

#### 方案 C：Claude Code / Codex / 通用 Agent

```bash
git clone https://github.com/drydraft/weekly-report-skill.git \
  ~/.agents/skills/weekly-report
```

### 2. 依赖检查

生成脚本基于 Python 与 `openpyxl`：

```bash
pip install openpyxl
```

---

## 💬 使用体验示例

只需在你的 Agent 会话中自然打字：

```text
用户：帮我写下本周周报，这周我排查了结算流水线的慢查询，另外完成了多租户权限接口开发，下周准备做订单网关回归测试。

AI：收到！已为你进行职场化提炼，请确认以下细节：
  1. 结算慢查询优化的完成状态是已完成吗？
  2. 多租户权限接口联调的协作人是谁？
  3. 下周订单网关回归测试的计划起止时间是？

用户：慢查询已完成；协作人是李四；下周一到周三做网关回归。

AI：📋【本周周报预览】
  - 周期：2026/05/25 - 2026/05/29
  - 姓名/部门：张三 (Alex Zhang) / 核心业务研发部
  --------------------------------------------------
  1. 上周完成情况：
     - [已完成] 排查并定位生产环境订单结算响应延迟瓶颈，优化慢查询 SQL 索引，结算响应耗时降低 40%
  2. 本周重点工作：
     - [高] 推进用户中心多租户权限 RESTful 接口开发与联调（协作人：李四）
  3. 下周工作计划：
     - [高] 推进异步消息通知网关核心业务流回归测试与发布验证（2026-06-01至2026-06-03）
  4. 饱和度评估：
     - 饱和度：饱和（双线并行推进，开发与性能优化进度符合预期）
  --------------------------------------------------
  👉 请确认以上内容无误，回复「确认」后我将为您写入企业标准 Excel 模板。

用户：确认

AI：🎉 周报已生成完毕！
    文件路径：C:\Users\Administrator\Desktop\周报\工作周报 - 张三 2026-05-25至2026-05-29.xlsx
```

---

## 🛠️ CLI 开发者接口

技能内部自带完备的命令行工具链（`scripts/generate_report.py`），支持自动化流水线调用与排障：

```bash
# 1. 验证模板完整性
python scripts/generate_report.py --validate-template

# 2. 从 JSON 载荷直接生成 Excel
python scripts/generate_report.py payload.json

# 3. 查看上一份周报的完整结构化数据
python scripts/generate_report.py --last-week

# 4. 提取上周周报的顺延迁移建议（JSON 输出）
python scripts/generate_report.py --migrate-last-week

# 5. 运行自检测试
python scripts/generate_report.py --test
```

---

## ⚙️ 配置与自定义选项

无需修改代码，可通过环境变量覆盖默认行为：

| 环境变量 | 说明 | 默认值 |
| --- | --- | --- |
| `WKR_TEMPLATE_PATH` | 自定义 Excel 模板路径 | Skill 自带 `templates/工作周报模板.xlsx` |
| `WKR_OUTPUT_DIR` | 周报 Excel 默认保存目录 | 用户桌面目录 `~/Desktop/周报/` |

---

## 📁 项目结构

```text
weekly-report-skill/
├── assets/
│   └── readme/
│       ├── hero.svg                 # 响应式项目封面大图
│       └── workflow.svg             # 执行与交互流程图
├── scripts/
│   └── generate_report.py           # 工业级 CLI 与 openpyxl 写入引擎
├── templates/
│   └── 工作周报模板.xlsx      # 开箱自带的标准 Excel 模板
├── tests/
│   └── test_generate_report.py      # 单元测试集（100% 覆盖关键路径）
├── SKILL.md                         # Agent 技能规范（控制流、安全锁、润色定义）
└── README.md                        # 本文档
```

---

## 🤝 贡献与反馈

欢迎提交 Issue 或 Pull Request！

* 发现字段映射异常或模板不兼容 $\rightarrow$ [提交 Issue](https://github.com/drydraft/weekly-report-skill/issues)
* 适配更多公司定制模板或功能增强 $\rightarrow$ 欢迎提 PR！

---

<p align="center">
  Made with ☕ by <a href="https://github.com/drydraft">drydraft</a> · Powered by <a href="https://agentskills.io">Agent Skills</a>
</p>
