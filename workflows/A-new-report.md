# 流程 A：全新周报填报

上级：[SKILL.md](../SKILL.md)。字段清单见 [fields.md](../references/fields.md)，payload 结构见 [cli-and-payload.md](../references/cli-and-payload.md)，润色标准见 [polishing.md](../references/polishing.md)。

适用于用户从零开始写一份新周报。支持两种入口，可混用：

- **向导式**：AI 逐板块问，适合用户给的信息很少。
- **自由式**：用户一次性贴一大段碎碎念，AI 自行拆解归档，只对缺的追问。

---

## 门 1 — 收集与润色

1. 先算周期：用户没给 → 按本周一至周五推；跨周或不确定就一句话确认。
2. 按 `references/fields.md` 的四大板块逐项对齐，**已给的直接采用，缺的一轮问完**（≤3 问）。
3. 每条内容过一遍 `references/polishing.md` 的润色公式后再入草稿。
4. 逐项确认时优先补这三类高频缺失：**协作人**、**预计完成时间 / eta**、**成果说明**。
5. 用户描述模糊 → 追问到具体系统/动作，不要硬猜。

**自由式入口的拆解要点**：

- 把用户的话按「已完成 / 在做 / 打算做」归到上周完成 / 本周重点 / 下周计划三栏。
- 用户提到的同事实名 → `collaborators`；不要漏。
- 用户**没给日期时不要编**，留空并在预览里标 `待确认`，让用户补。

**Postconditions**：四大板块字段齐全（含每条协作人）、周期与姓名部门确定。

## 门 2 — 预览与确认

按下面的模板输出 Markdown 预览（数字、状态、协作人要与最终 payload 完全一致）：

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

- 用 `AskUserQuestion` 或请用户直接回复确认。
- 用户提修改 → 改 payload 草稿 → **重新出预览**，不得带着未确认的改动进入门 3。
- 若用了 `fields.md` 的默认值（姓名、部门、饱和度），在预览里显式标出来。

**Postconditions**：用户明确确认。

## 门 3 — 校验 → 生成 → 交付

**数据流协议（严格按序）**：

1. 用 `write` 把 payload 写到工作目录 `.weekly_report_payload.json`。
2. 校验：

   ```powershell
   python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\validate_payload.py" .weekly_report_payload.json
   ```

   期望 `{"status":"ok",...}`。返回 `invalid` → 按 `errors` 修正后重跑（最多 3 轮）。
3. 生成：

   ```powershell
   python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\generate_report.py" .weekly_report_payload.json
   ```

   期望 `{"status":"success","file":"C:\\...\\工作周报 - 张三 2026-05-25至2026-05-29.xlsx"}`。
4. 删除临时 JSON，用 `present_files` 打开生成的文件（或给出绝对路径），并回报：

   ```
   🎉 周报已生成：<绝对路径>
   校验：validate_payload 通过 · 生成：success
   ```

**Postconditions**：脚本返回 success 且文件存在。

> ⛔ 任一步失败都不要谎报成功。`status: error` 就把 `message` 原样告诉用户并给出下一步动作。
