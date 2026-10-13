# 流程 B：基于上周周报顺延更新

上级：[SKILL.md](../SKILL.md)。触发词：「基于上周周报写」「更新周报」「接着上周写」「顺延」。

## 1. 取顺延建议

```powershell
python "$env:USERPROFILE\.pi\agent\skills\weekly-report\scripts\generate_report.py" --migrate-last-week
```

脚本扫描输出目录里最新的一份周报（`*周报 - *.xlsx`），返回：

| 返回字段 | 来源 | 用途 |
| --- | --- | --- |
| `suggested_last_week` | 上上周报的**本周重点** | 建议作为本次的「上周完成情况」 |
| `suggested_this_week` | 上上周报的**下周计划** | 建议作为本次的「本周重点工作」 |

需要完整原文（不只要建议）时改用 `--last-week`。

## 2. 与用户核对差量

把建议整理成一段话给用户，**只问缺的**：

- 上周承诺的事**实际做到没有**？→ 定 `status`（`已完成` / `进行中` / `延期` / `未开始`）与 `results`。
- 本周重点的**优先级、协作人、eta** 是否变化？→ 变了的改，没变的沿用。
- 下周计划 → 由用户口头补充，或沿用上上周报的下周计划里未完成的部分。

不得把建议当成事实直接填 —— 状态必须由用户确认。

## 3. 组装新周期 payload

在新 payload 里：

- `period` 换成**本周**周期（不是上一次的）。
- `last_week` ← 已确认的建议 + 状态修正。
- `this_week` ← 已确认的顺延项 + 用户新增项。
- `next_week` ← 用户补充项。
- `workload_desc` / `resources_needed` 按本周实际情况重写，**不要照抄上一周的文字**。

## 4. 预览 → 校验 → 生成

与 [A-new-report.md](A-new-report.md) 的门 2、门 3 完全一致：

1. 出 Markdown 预览卡片 → 拿到用户明确确认。
2. 写 `.weekly_report_payload.json`。
3. `validate_payload.py` 通过。
4. `generate_report.py` 生成并交付。

> 建立在**上周文件已被移动或删除**的前提下：`read_last_week_report()` 按修改时间取输出目录里最新的 `*周报 - *.xlsx`，若取到的不是上周那份，先用 `--last-week` 让用户核对内容，确认对了再继续。
