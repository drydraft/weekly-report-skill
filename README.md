<p align="center">
  <img src="./assets/readme/hero.svg" width="100%" alt="周报助手 Weekly Report Assistant：对 AI 讲几句大白话碎碎念，自动完成职场化提炼，并填充至公司标准 Excel 周报模板。">
</p>

<p align="center">
  <a href="https://agentskills.io"><img src="https://img.shields.io/badge/standard-Agent_Skills-0284C7?style=flat-square" alt="Agent Skills Standard"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-10B981?style=flat-square" alt="MIT License"></a>
  <a href="#-支持平台"><img src="https://img.shields.io/badge/runtime-Pi_Agent_|_OpenCode_|_Claude_Code-8B5CF6?style=flat-square" alt="Runtimes"></a>
  <a href="#-为什么需要周报助手"><img src="https://img.shields.io/badge/Excel-100%25_Fidelity-107C41?style=flat-square" alt="Excel Fidelity"></a>
  <a href="#-校验与测试"><img src="https://img.shields.io/badge/checks-22_tests_passing-0891B2?style=flat-square" alt="Tests"></a>
</p>

---

## 💡 为什么需要周报助手？

每到周五下午，写周报往往面临两大痛点：

1. **话术包装难**：做了一堆事情，但写出来全是"改了bug"、"对了接口"，口语化太重，难以体现工作成果与业务价值；
2. **填表排版烦**：打开公司动辄几十列、带合并单元格和公式的 Excel 模板，复制粘贴经常搞坏样式、漏填协作人。

**周报助手（Weekly Report Skill）** 把这个流程拆成两半 —— **把内容提炼交给大模型，把文件落地交给确定性 Python 引擎**。你只需说几句本周做的大白话，AI 完成职场规范化重构，零损保真地填充进公司标准 `.xlsx` 模板，直达桌面。

<p align="center">
  <img src="./assets/readme/workflow.svg" width="100%" alt="周报助手工作流：自然语言碎碎念 → 智能职场化提炼 → 预览卡片与确认锁 → 载荷校验 → 原生 Excel 安全写入。">
</p>

---

## ✨ 核心特性

### 1. 智能职场化润色（底层核心能力）

不是机械的关键词替换，而是理解业务上下文后，按 **【精准动词】+【业务系统/模块】+【关键手段/重心】+【阶段价值】** 结构化重塑：

| 原始口语输入 | 职场规范化润色呈现（进 Excel 效果） |
| --- | --- |
| "改了结算流水线的并发bug" | **排查并修复订单结算流水线并发冲突异常，优化分布式锁边界校验，确保交易幂等无误** |
| "跟产品对了多租户权限二期" | **参与用户中心多租户权限系统业务需求评审，对齐核心角色鉴权规则与验收标准，完成技术任务拆解** |
| "写了几个接口并自测" | **完成核心业务模块 RESTful 接口开发与联调，通过接口幂等性验证与单元测试覆盖** |
| "线上有点卡，查了下" | **排查并定位生产环境服务响应延迟瓶颈，优化慢查询 SQL 索引与缓存策略，提升接口响应效率** |

润色只改**表达精度**，不改事实：不允许编造未发生的功能、虚构数据指标或捏造外部协作（见 [`references/polishing.md`](references/polishing.md)）。

### 2. 100% 原生 Excel 零损保真

- 拒绝"生成一张不可选中文本的大长图"，也拒绝"导出一个丢了样式的 csv"。
- 用 `openpyxl` 在**单元格级别**操作原生 `.xlsx`，完整保留企业模板的表头配色、合并单元格、进度计算公式与边框。
- 含换行的文本自动套用 `wrap_text` + 顶端对齐。

### 3. 跨周智能顺延继承

用户只要说"**基于上周周报写**"：

- `--migrate-last-week` 自动解析上一份周报，把上上周报的「本周重点」顺延为「上周完成」、把「下周计划」顺延为「本周重点」；
- AI 只跟用户核对差量（实际状态、变动的人与日期），大幅省掉重复输入。

### 4. 三道门的安全锁

| 门 | 内容 | 不通过会怎样 |
| --- | --- | --- |
| 门 1 · 收集与润色 | 四大板块字段齐全（含每条协作人） | 停在收集，最多追问 2 轮，其余用默认值并显式告知 |
| 门 2 · 预览与确认 | 必须呈递 Markdown 预览卡片并拿到用户明确确认 | 绝不生成文件 |
| 门 3 · 校验与生成 | 先过载荷校验器，再交给生成引擎 | 校验不过就不写盘 |

**绝不提前生成文件**、**严禁用文本工具直接改 `.xlsx`** —— 这两条写进了 `SKILL.md` 的硬门禁。

### 5. 载荷校验器（机器可验的通过条件）

写盘之前先跑 `validate_payload.py`，把"协作人漏填、优先级写成『紧急』、日期写成『下周三』、条数超过模板 5 行上限、周期只给了一个日期"这类问题挡在前面，并给出**可定位到字段路径**的报错：

```json
{
  "status": "invalid",
  "errors": [
    "this_week[0]: 协作人未填 —— 铁律要求必须主动询问，不能留空带队",
    "last_week[0]: status 只能取 已完成/进行中/未开始/延期，当前 '搞定'",
    "next_week[1]: 缺少必填字段 eta"
  ]
}
```

---

## 🧭 Skill 架构

本 Skill 按 **"渐进披露 + 硬门禁 + 机器校验"** 组织：`SKILL.md` 只做路由与控制流，知识全部按需加载。

```text
SKILL.md                      # 控制流：三道门、7 条硬门禁、Execution Protocol
├─ references/                # 用到才读的知识
│  ├─ fields.md               #   四大板块必填清单 · 取值枚举 · 字段→单元格映射
│  ├─ polishing.md            #   润色公式 · 转换对照表 · 真实性底线
│  └─ cli-and-payload.md      #   CLI 命令表 · Payload 结构 · 环境变量
├─ workflows/                 # 端到端流程
│  ├─ A-new-report.md         #   全新填报（含预览卡片模板与数据流协议）
│  └─ B-migrate-last-week.md  #   基于上周顺延更新
└─ scripts/                   # 直接执行，不读进上下文
   ├─ generate_report.py      #   生成引擎（openpyxl 写入）
   └─ validate_payload.py     #   写盘前的静态校验器（无第三方依赖）
```

三条设计原则：

1. **SKILL.md 不装知识，只装"什么情况读哪个文件"** —— 主文件保持精简，长尾细节不进上下文。
2. **每一步都有机器可验的通过条件** —— 校验器报错不带"感觉不对"，而是精确到 `this_week[0].collaborators`。
3. **收敛阈值** —— 追问最多 2 轮、校验修正最多 3 轮、成功后不做主观重排，避免无限自我修正。

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

生成引擎基于 Python 与 `openpyxl`（校验器只用标准库，未装 `openpyxl` 时也能跑）：

```bash
pip install openpyxl
python scripts/generate_report.py --validate-template   # 应输出 status: success
```

---

## 💬 使用体验示例

只需在你的 Agent 会话里自然打字：

```text
用户：帮我写下本周周报，这周我排查了结算流水线的慢查询，另外完成了多租户权限接口开发，下周准备做订单网关回归测试。

AI：收到！已为你进行职场化提炼，请确认以下细节：
  1. 结算慢查询优化的完成状态是已完成吗？
  2. 多租户权限接口联调的协作人是谁？
  3. 下周订单网关回归测试的计划起止时间是？

用户：慢查询已完成；协作人是李四；下周一到周三做网关回归。

AI：[门 2] 预览与确认 — preconditions: PASS
  📋 周报预览
  - 周期：2026/05/25 - 2026/05/29
  - 姓名/部门：张三 / 核心业务研发部
  --------------------------------------------------
  1. 上周完成情况：
     - [已完成] 排查并定位生产环境订单结算响应延迟瓶颈，优化慢查询 SQL 索引
  2. 本周重点工作：
     - [高] 推进用户中心多租户权限 RESTful 接口开发与联调（协作人：李四）
  3. 下周工作计划：
     - [高] 推进异步消息通知网关核心业务流回归测试与发布验证（2026-06-01至2026-06-03）
  4. 饱和度评估：
     - 饱和度：饱和（双线并行推进，开发与性能优化进度符合预期）
  --------------------------------------------------
  👉 请确认以上内容无误，回复「确认」后我将为您写入企业标准 Excel 模板。

用户：确认

AI：[门 3] 校验 → 生成 → 交付 — preconditions: PASS
  校验：validate_payload → ok
  生成：generate_report  → success
  🎉 文件路径：C:\Users\Administrator\Desktop\周报\工作周报 - 张三 2026-05-25至2026-05-29.xlsx
```

---

## 🛠️ CLI 接口

技能自带完备的命令行工具链，可用于自动化流水线与排障：

```bash
# 0. 写盘前校验载荷（推荐先跑）
python scripts/validate_payload.py payload.json

# 1. 验证模板完整性（会回显识别到的区块行）
python scripts/generate_report.py --validate-template

# 2. 从 JSON 载荷生成 Excel
python scripts/generate_report.py payload.json

# 3. 读取上一份周报的完整结构化数据
python scripts/generate_report.py --last-week

# 4. 提取上周周报的顺延迁移建议
python scripts/generate_report.py --migrate-last-week

# 5. 运行内置示例生成自检
python scripts/generate_report.py --test
```

`generate_report.py` 也支持从 stdin 读（`-`）与内联 JSON（`--json '<json>'`）。载荷结构见 [`references/cli-and-payload.md`](references/cli-and-payload.md)。

---

## ⚙️ 配置与自定义

无需改代码，用环境变量覆盖默认行为：

| 环境变量 | 说明 | 默认值 |
| --- | --- | --- |
| `WKR_TEMPLATE_PATH` | 自定义 Excel 模板路径 | Skill 自带 `templates/工作周报模板.xlsx` |
| `WKR_OUTPUT_DIR` | 周报 Excel 默认保存目录 | 用户桌面目录 `~/Desktop/周报/` |

换成自家公司模板时：把模板放到 `templates/` 或设 `WKR_TEMPLATE_PATH` 指向它，再跑一次 `--validate-template` 确认脚本能识别到四个区块（`上周工作完成情况` / `本周重要工作项` / `下周工作计划` / `工作饱和度评估`）。区块是按 **A 列表头文字**动态定位的，所以改版模板只要保留这些标题即可。

---

## ✅ 校验与测试

```bash
# 单元测试（22 个用例：生成引擎 + 校验器）
python -m pytest tests/ -q

# 生成引擎自带自检
python scripts/generate_report.py --test
python scripts/generate_report.py --validate-template

# 校验器单独试跑
python scripts/validate_payload.py --json '{"period":"2026/05/25-2026/05/29","name":"张三","department":"IT部","last_week":[],"this_week":[],"next_week":[]}'
```

行为契约由 `evals/evals.json` 描述：每份周报都必须过三道门，且饱和度固定写「饱和」、文件名固定为 `工作周报 - {姓名} {起始}至{结束}.xlsx`。

---

## 📁 项目结构

```text
weekly-report-skill/
├── SKILL.md                        # Agent 技能规范：控制流 · 硬门禁 · Execution Protocol
├── README.md                       # 本文档
├── LICENSE                         # MIT
├── references/                     # 按需加载的知识
│   ├── fields.md                   #   必填字段 · 取值枚举 · 单元格映射
│   ├── polishing.md                #   职场化润色规范
│   └── cli-and-payload.md          #   CLI · Payload · 环境变量
├── workflows/                      # 端到端流程
│   ├── A-new-report.md             #   全新填报
│   └── B-migrate-last-week.md      #   顺延更新
├── scripts/
│   ├── generate_report.py          # 生成引擎（openpyxl 写入，自带 CLI）
│   └── validate_payload.py         # 载荷静态校验器（纯标准库）
├── templates/
│   └── 工作周报模板.xlsx           # 开箱自带的标准 Excel 模板
├── tests/
│   ├── test_generate_report.py     # 生成引擎单测
│   └── test_validate_payload.py    # 校验器单测
├── evals/
│   └── evals.json                  # 行为验收用例
└── assets/
    └── readme/
        ├── hero.svg                # 响应式项目封面大图
        └── workflow.svg            # 执行与交互流程图
```

---

## 🤝 贡献与反馈

欢迎提交 Issue 或 Pull Request！

- 发现字段映射异常或模板不兼容 → [提交 Issue](https://github.com/drydraft/weekly-report-skill/issues)
- 适配更多公司定制模板或功能增强 → 欢迎提 PR！提交前请跑一遍 `python -m pytest tests/ -q`

---

<p align="center">
  Made with ☕ by <a href="https://github.com/drydraft">drydraft</a> · Powered by <a href="https://agentskills.io">Agent Skills</a>
</p>
